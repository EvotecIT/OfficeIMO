using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    internal static partial class WordFieldInventory {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private static readonly Regex FieldTypeTokenPattern = new Regex(
            @"^\s*(?<field>[A-Za-z]+)",
            RegexOptions.CultureInvariant);

        private static readonly Regex SupportedNumericPictureSwitchPattern = new Regex(
            @"^(?<instruction>.*?)\s+\\#\s*(?<format>""[^""]*""|[^\s\\]+)(?<suffix>.*)$",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);

        internal static IReadOnlyList<WordFieldInfo> Inspect(WordDocument document) {
            var findings = InspectStories(EnumerateStories(document));
            return findings.Select(finding => finding.Info).ToArray();
        }

        private static FieldInfoCandidate[] InspectStories(IEnumerable<NodeFieldRoot> stories) {
            var findings = new List<FieldInfoCandidate>();
            foreach (NodeFieldRoot story in stories) InspectStory(story, findings);
            return findings.ToArray();
        }

        private static void InspectStory(NodeFieldRoot story, List<FieldInfoCandidate> findings) {
            var stack = new Stack<ComplexFieldBuilder>();
            foreach (OfficeXmlNode element in EnumerateFieldContent(story.Root)) {
                if (element.Is(W, "fldSimple")) {
                    findings.Add(ReadSimpleField(story, element, stack.Count, findings.Count));
                } else if (element.Is(W, "r")) {
                    ProcessRun(story, element, stack, findings);
                }
            }

            while (stack.Count > 0) {
                ComplexFieldBuilder builder = stack.Pop();
                builder.Diagnostics.Add("Complex field does not have an end marker.");
                findings[builder.Index] = builder.ToCandidate(story);
            }
        }

        private static FieldInfoCandidate ReadSimpleField(NodeFieldRoot story, OfficeXmlNode field, int complexNestingLevel, int index) {
            string instructionText = field.Attribute(W, "instr") ?? string.Empty;
            ParsedFieldInstruction parsed = ParseInstruction(instructionText);
            int nestingLevel = complexNestingLevel;
            for (OfficeXmlNode ancestor = field.Parent; ancestor.Exists; ancestor = ancestor.Parent) {
                if (ancestor.Is(W, "fldSimple")) nestingLevel++;
            }

            var result = new StringBuilder();
            bool hasStoredResult = false;
            foreach (OfficeXmlNode child in EnumerateFieldContent(field)) {
                if (child.Is(W, "t")) result.Append(child.InnerText);
                if (child.Is(W, "r") || child.Is(W, "contentPart") || child.Is(W, "subDoc") ||
                    child.NamespaceId == OfficeXmlNamespaces.Math) hasStoredResult = true;
            }

            var info = new WordFieldInfo(
                index: index,
                representation: WordFieldRepresentation.Simple,
                locationKind: story.LocationKind,
                partUri: story.PartUri,
                instructionText: instructionText,
                resultText: result.ToString(),
                fieldType: parsed.FieldType,
                instructions: parsed.Instructions,
                switches: parsed.Switches,
                formatSwitches: parsed.FormatSwitches,
                isDirty: ReadFlag(field, "dirty"),
                isLocked: ReadFlag(field, "fldLock"),
                nestingLevel: nestingLevel,
                isInTable: field.Parent.AncestorOrSelf(W, "tbl").Exists,
                isInContentControl: field.Parent.AncestorOrSelf(W, "sdt").Exists,
                isInTextBox: IsInTextBox(field),
                unsupportedParseDetails: parsed.Diagnostics);

            return new FieldInfoCandidate(info, field, hasStoredResult);
        }

        private static void ProcessRun(NodeFieldRoot story, OfficeXmlNode run, Stack<ComplexFieldBuilder> stack,
            List<FieldInfoCandidate> findings) {
            foreach (OfficeXmlNode child in run.Children()) {
                if (child.Is(W, "fldChar")) {
                    string? fieldCharType = child.Attribute(W, "fldCharType");
                    if (fieldCharType == "begin") {
                        // Reserve the position at the begin marker, before any nested fields.
                        stack.Push(new ComplexFieldBuilder(findings.Count, stack.Count, child) {
                            IsDirty = ReadFlag(child, "dirty"),
                            IsLocked = ReadFlag(child, "fldLock"),
                            IsInTable = run.Parent.AncestorOrSelf(W, "tbl").Exists,
                            IsInContentControl = run.Parent.AncestorOrSelf(W, "sdt").Exists,
                            IsInTextBox = IsInTextBox(run)
                        });
                        findings.Add(null!);
                        continue;
                    }

                    if (stack.Count == 0) continue;

                    if (fieldCharType == "separate") {
                        stack.Peek().HasSeparator = true;
                    } else if (fieldCharType == "end") {
                        ComplexFieldBuilder completed = stack.Pop();
                        if (stack.Count > 0 && !stack.Peek().HasSeparator) {
                            stack.Peek().InstructionParts.Add(completed.ResultText);
                        }

                        findings[completed.Index] = completed.ToCandidate(story);
                    }
                    continue;
                }

                if (stack.Count == 0) continue;

                if (child.Is(W, "instrText")) {
                    stack.Peek().InstructionParts.Add(child.InnerText);
                } else if (child.Is(W, "t")) {
                    string text = child.InnerText;
                    foreach (ComplexFieldBuilder builder in stack) {
                        if (builder.HasSeparator) builder.ResultParts.Add(text);
                    }
                }
            }
        }

        internal static bool ReadFlag(OfficeXmlNode field, string attribute) =>
            field.Attribute(W, attribute) is "true" or "1" or "on";

        private static bool IsInTextBox(OfficeXmlNode element) {
            for (OfficeXmlNode ancestor = element.Parent; ancestor.Exists; ancestor = ancestor.Parent) {
                if (ancestor.LocalName is "txbxContent" or "textbox") return true;
            }
            return false;
        }

        internal static IEnumerable<NodeFieldRoot> EnumerateStories(WordDocument document, bool includeNotes = true) =>
            WordPackageView.For(document).Main is { } main ? EnumerateFieldRoots(main, includeNotes) : Enumerable.Empty<NodeFieldRoot>();

        /// <summary>The body, headers, footers, footnotes and endnotes of a package, in the order fields are inspected.</summary>
        internal static IEnumerable<NodeFieldRoot> EnumerateFieldRoots(WordPartView mainPart, bool includeNotes = true) {
            const int W = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.WordprocessingML;
            OfficeXmlNode body = mainPart.RootNamed(W, "document").Child(W, "body");
            if (!body.IsNone) {
                yield return new NodeFieldRoot(body, WordFieldLocationKind.Body, mainPart.PartName, mainPart);
            }

            foreach (WordPartView headerPart in mainPart.Parts(WordPackageView.HeaderRelationshipType)
                .OrderBy(part => part.PartName, StringComparer.Ordinal)) {
                OfficeXmlNode header = headerPart.RootNamed(W, "hdr");
                if (!header.IsNone) {
                    yield return new NodeFieldRoot(header, WordFieldLocationKind.Header, headerPart.PartName, headerPart);
                }
            }

            foreach (WordPartView footerPart in mainPart.Parts(WordPackageView.FooterRelationshipType)
                .OrderBy(part => part.PartName, StringComparer.Ordinal)) {
                OfficeXmlNode footer = footerPart.RootNamed(W, "ftr");
                if (!footer.IsNone) {
                    yield return new NodeFieldRoot(footer, WordFieldLocationKind.Footer, footerPart.PartName, footerPart);
                }
            }

            if (!includeNotes) yield break;

            WordPartView? footnotesPart = mainPart.FirstPart(WordPackageView.FootnotesRelationshipType);
            OfficeXmlNode footnotes = footnotesPart?.RootNamed(W, "footnotes") ?? default;
            if (!footnotes.IsNone) {
                yield return new NodeFieldRoot(footnotes, WordFieldLocationKind.Footnote, footnotesPart!.PartName, footnotesPart);
            }

            WordPartView? endnotesPart = mainPart.FirstPart(WordPackageView.EndnotesRelationshipType);
            OfficeXmlNode endnotes = endnotesPart?.RootNamed(W, "endnotes") ?? default;
            if (!endnotes.IsNone) {
                yield return new NodeFieldRoot(endnotes, WordFieldLocationKind.Endnote, endnotesPart!.PartName, endnotesPart);
            }
        }

        /// <summary>A field root read through <see cref="WordPartView"/>, with the part that holds it.</summary>
        internal sealed class NodeFieldRoot {
            internal NodeFieldRoot(OfficeXmlNode root, WordFieldLocationKind locationKind, string partUri, WordPartView part) {
                Root = root;
                LocationKind = locationKind;
                PartUri = partUri;
                Part = part;
            }

            internal OfficeXmlNode Root { get; }

            internal WordFieldLocationKind LocationKind { get; }

            internal string PartUri { get; }

            internal WordPartView Part { get; }
        }

        internal static ParsedFieldInstruction ParseInstruction(string instructionText) {
            if (string.IsNullOrWhiteSpace(instructionText)) {
                return new ParsedFieldInstruction(
                    null,
                    Array.Empty<string>(),
                    Array.Empty<string>(),
                    Array.Empty<WordFieldFormat>(),
                    null,
                    new[] { "Field instruction is empty." });
            }

            string trimmedInstruction = instructionText.Trim();
            if (trimmedInstruction.StartsWith("=", StringComparison.Ordinal)) {
                return new ParsedFieldInstruction(
                    WordFieldType.Formula,
                    new[] { trimmedInstruction.Substring(1).Trim() },
                    Array.Empty<string>(),
                    Array.Empty<WordFieldFormat>(),
                    null,
                    Array.Empty<string>());
            }

            try {
                if (!TryExtractSupportedNumericPictureSwitch(instructionText, out string parserInstruction, out string? numericPicture, out WordFieldType? parsedFieldType, out string? diagnostic)) {
                    return new ParsedFieldInstruction(
                        parsedFieldType,
                        Array.Empty<string>(),
                        Array.Empty<string>(),
                        Array.Empty<WordFieldFormat>(),
                        numericPicture,
                        new[] { diagnostic ?? "Field numeric picture switch could not be parsed." });
                }

                var parser = new WordFieldParser(parserInstruction);
                return new ParsedFieldInstruction(
                    parser.WordFieldType,
                    parser.Instructions.ToArray(),
                    parser.Switches.ToArray(),
                    parser.FormatSwitches.ToArray(),
                    numericPicture,
                    parser.Diagnostics.ToArray());
            } catch (Exception ex) when (ex is NotImplementedException || ex is ArgumentException || ex is InvalidOperationException) {
                return new ParsedFieldInstruction(
                    null,
                    Array.Empty<string>(),
                    Array.Empty<string>(),
                    Array.Empty<WordFieldFormat>(),
                    null,
                    new[] { ex.Message });
            }
        }

        private static bool TryExtractSupportedNumericPictureSwitch(
            string instructionText,
            out string instructionWithoutPicture,
            out string? numericPicture,
            out WordFieldType? fieldType,
            out string? diagnostic) {
            instructionWithoutPicture = instructionText;
            numericPicture = null;
            fieldType = null;
            diagnostic = null;

            if (!TryGetFieldType(instructionText, out WordFieldType detectedFieldType)) {
                return true;
            }

            fieldType = detectedFieldType;
            if (detectedFieldType != WordFieldType.Page &&
                detectedFieldType != WordFieldType.FileSize &&
                detectedFieldType != WordFieldType.NumPages &&
                detectedFieldType != WordFieldType.NumWords &&
                detectedFieldType != WordFieldType.NumChars &&
                detectedFieldType != WordFieldType.Section &&
                detectedFieldType != WordFieldType.SectionPages &&
                detectedFieldType != WordFieldType.PageRef &&
                detectedFieldType != WordFieldType.Quote &&
                detectedFieldType != WordFieldType.MergeField) {
                return true;
            }

            if (instructionText.IndexOf(@"\#", StringComparison.Ordinal) < 0) {
                return true;
            }

            Match match = SupportedNumericPictureSwitchPattern.Match(instructionText);
            if (!match.Success) {
                diagnostic = $"{detectedFieldType} field numeric picture switch must appear at the end of the field instruction.";
                return false;
            }

            instructionWithoutPicture = (match.Groups["instruction"].Value + match.Groups["suffix"].Value).Trim();
            numericPicture = match.Groups["format"].Value.Trim();
            return true;
        }

        private static bool TryGetFieldType(string instructionText, out WordFieldType fieldType) {
            fieldType = default;
            Match match = FieldTypeTokenPattern.Match(instructionText);
            return match.Success &&
                Enum.TryParse(match.Groups["field"].Value, ignoreCase: true, out fieldType);
        }

        internal sealed partial class FieldInfoCandidate {
            private readonly OfficeXmlNode _source;

            internal FieldInfoCandidate(WordFieldInfo info, OfficeXmlNode source, bool hasStoredResult) {
                Info = info;
                _source = source;
                HasStoredResult = hasStoredResult;
            }

            internal WordFieldInfo Info { get; }
            internal bool HasStoredResult { get; }
        }

        private sealed class ComplexFieldBuilder {
            internal ComplexFieldBuilder(int index, int nestingLevel, OfficeXmlNode startElement) {
                Index = index;
                NestingLevel = nestingLevel;
                StartElement = startElement;
            }

            internal int Index { get; }

            internal int NestingLevel { get; }

            internal OfficeXmlNode StartElement { get; }

            internal List<string> InstructionParts { get; } = new();

            internal List<string> ResultParts { get; } = new();

            internal List<string> Diagnostics { get; } = new();

            internal bool HasSeparator { get; set; }

            internal bool IsDirty { get; set; }

            internal bool IsLocked { get; set; }

            internal bool IsInTable { get; set; }

            internal bool IsInContentControl { get; set; }

            internal bool IsInTextBox { get; set; }

            internal string ResultText => string.Concat(ResultParts);

            internal FieldInfoCandidate ToCandidate(NodeFieldRoot story) {
                string instructionText = string.Concat(InstructionParts);
                ParsedFieldInstruction parsed = ParseInstruction(instructionText);
                var diagnostics = Diagnostics.Concat(parsed.Diagnostics).ToArray();

                var info = new WordFieldInfo(
                    index: Index,
                    representation: WordFieldRepresentation.Complex,
                    locationKind: story.LocationKind,
                    partUri: story.PartUri,
                    instructionText: instructionText,
                    resultText: ResultText,
                    fieldType: parsed.FieldType,
                    instructions: parsed.Instructions,
                    switches: parsed.Switches,
                    formatSwitches: parsed.FormatSwitches,
                    isDirty: IsDirty,
                    isLocked: IsLocked,
                    nestingLevel: NestingLevel,
                    isInTable: IsInTable,
                    isInContentControl: IsInContentControl,
                    isInTextBox: IsInTextBox,
                    unsupportedParseDetails: diagnostics);

                return new FieldInfoCandidate(info, StartElement, HasSeparator);
            }
        }

        internal sealed class ParsedFieldInstruction {
            internal ParsedFieldInstruction(
                WordFieldType? fieldType,
                IReadOnlyList<string> instructions,
                IReadOnlyList<string> switches,
                IReadOnlyList<WordFieldFormat> formatSwitches,
                string? numericPictureSwitch,
                IReadOnlyList<string> diagnostics) {
                FieldType = fieldType;
                Instructions = instructions;
                Switches = switches;
                FormatSwitches = formatSwitches;
                NumericPictureSwitch = numericPictureSwitch;
                Diagnostics = diagnostics;
            }

            internal WordFieldType? FieldType { get; }

            internal IReadOnlyList<string> Instructions { get; }

            internal IReadOnlyList<string> Switches { get; }

            internal IReadOnlyList<WordFieldFormat> FormatSwitches { get; }

            internal string? NumericPictureSwitch { get; }

            internal IReadOnlyList<string> Diagnostics { get; }
        }
    }
}
