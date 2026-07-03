using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    internal static class WordFieldInventory {
        private static readonly Regex FieldTypeTokenPattern = new Regex(
            @"^\s*(?<field>[A-Za-z]+)",
            RegexOptions.CultureInvariant);

        private static readonly Regex SupportedNumericPictureSwitchPattern = new Regex(
            @"^(?<instruction>.*?)\s+\\#\s*(?<format>""[^""]*""|[^\s\\]+)(?<suffix>.*)$",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);

        internal static IReadOnlyList<WordFieldInfo> Inspect(WordDocument document) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");

            var findings = new List<FieldInfoCandidate>();
            int sequence = 0;

            foreach (FieldRoot root in EnumerateFieldRoots(mainPart)) {
                InspectRoot(root, findings, ref sequence);
            }

            var ordered = findings
                .OrderBy(finding => finding.Sequence)
                .Select((finding, index) => finding.InfoWithIndex(index))
                .ToArray();

            return ordered;
        }

        private static void InspectRoot(FieldRoot root, List<FieldInfoCandidate> findings, ref int sequence) {
            var stack = new Stack<ComplexFieldBuilder>();

            foreach (OpenXmlElement element in root.Root.Descendants()) {
                if (element is SimpleField simpleField) {
                    AddSimpleField(root, simpleField, stack.Count, findings, ref sequence);
                    continue;
                }

                if (element is Run run) {
                    ProcessRun(root, run, stack, findings, ref sequence);
                }
            }

            while (stack.Count > 0) {
                ComplexFieldBuilder builder = stack.Pop();
                builder.Diagnostics.Add("Complex field does not have an end marker.");
                findings.Add(builder.ToCandidate(root));
            }
        }

        private static void AddSimpleField(
            FieldRoot root,
            SimpleField simpleField,
            int complexNestingLevel,
            List<FieldInfoCandidate> findings,
            ref int sequence) {
            string instructionText = simpleField.Instruction?.Value ?? string.Empty;
            ParsedFieldInstruction parsed = ParseInstruction(instructionText);
            int nestingLevel = complexNestingLevel + simpleField.Ancestors<SimpleField>().Count();
            bool isDirty = simpleField.Dirty?.Value ?? false;
            bool isLocked = simpleField.FieldLock?.Value ?? false;
            string resultText = string.Concat(simpleField.Descendants<Text>().Select(text => text.Text));

            var info = new WordFieldInfo(
                index: 0,
                representation: WordFieldRepresentation.Simple,
                locationKind: root.LocationKind,
                partUri: root.PartUri,
                instructionText: instructionText,
                resultText: resultText,
                fieldType: parsed.FieldType,
                instructions: parsed.Instructions,
                switches: parsed.Switches,
                formatSwitches: parsed.FormatSwitches,
                isDirty: isDirty,
                isLocked: isLocked,
                nestingLevel: nestingLevel,
                isInTable: IsInTable(simpleField),
                isInContentControl: IsInContentControl(simpleField),
                isInTextBox: IsInTextBox(simpleField),
                unsupportedParseDetails: parsed.Diagnostics);

            findings.Add(new FieldInfoCandidate(sequence++, info));
        }

        private static void ProcessRun(
            FieldRoot root,
            Run run,
            Stack<ComplexFieldBuilder> stack,
            List<FieldInfoCandidate> findings,
            ref int sequence) {
            foreach (OpenXmlElement child in run.ChildElements) {
                if (child is FieldChar fieldChar) {
                    FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                    if (fieldCharType == FieldCharValues.Begin) {
                        var builder = new ComplexFieldBuilder(sequence++, stack.Count, run) {
                            IsDirty = fieldChar.Dirty?.Value ?? false,
                            IsLocked = fieldChar.FieldLock?.Value ?? false,
                            IsInTable = IsInTable(run),
                            IsInContentControl = IsInContentControl(run),
                            IsInTextBox = IsInTextBox(run)
                        };

                        stack.Push(builder);
                        continue;
                    }

                    if (stack.Count == 0) {
                        continue;
                    }

                    if (fieldCharType == FieldCharValues.Separate) {
                        stack.Peek().HasSeparator = true;
                        continue;
                    }

                    if (fieldCharType == FieldCharValues.End) {
                        ComplexFieldBuilder completed = stack.Pop();
                        if (stack.Count > 0 && !stack.Peek().HasSeparator) {
                            stack.Peek().InstructionParts.Add(completed.ResultText);
                        }

                        findings.Add(completed.ToCandidate(root));
                    }

                    continue;
                }

                if (stack.Count == 0) {
                    continue;
                }

                if (child is FieldCode fieldCode) {
                    stack.Peek().InstructionParts.Add(fieldCode.Text ?? string.Empty);
                    continue;
                }

                if (child is Text runText) {
                    foreach (ComplexFieldBuilder builder in stack.Where(builder => builder.HasSeparator)) {
                        builder.ResultParts.Add(runText.Text);
                    }
                }
            }
        }

        internal static IEnumerable<FieldRoot> EnumerateFieldRoots(MainDocumentPart mainPart) {
            if (mainPart.Document?.Body != null) {
                yield return new FieldRoot(mainPart.Document.Body, WordFieldLocationKind.Body, mainPart.Uri.ToString());
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)) {
                if (headerPart.Header != null) {
                    yield return new FieldRoot(headerPart.Header, WordFieldLocationKind.Header, headerPart.Uri.ToString());
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)) {
                if (footerPart.Footer != null) {
                    yield return new FieldRoot(footerPart.Footer, WordFieldLocationKind.Footer, footerPart.Uri.ToString());
                }
            }

            FootnotesPart? footnotesPart = mainPart.FootnotesPart;
            if (footnotesPart?.Footnotes != null) {
                yield return new FieldRoot(footnotesPart.Footnotes, WordFieldLocationKind.Footnote, footnotesPart.Uri.ToString());
            }

            EndnotesPart? endnotesPart = mainPart.EndnotesPart;
            if (endnotesPart?.Endnotes != null) {
                yield return new FieldRoot(endnotesPart.Endnotes, WordFieldLocationKind.Endnote, endnotesPart.Uri.ToString());
            }
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
                detectedFieldType != WordFieldType.Quote) {
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

        private static bool IsInTable(OpenXmlElement element) => element.Ancestors<Table>().Any();

        private static bool IsInContentControl(OpenXmlElement element) => element.Ancestors<SdtElement>().Any();

        private static bool IsInTextBox(OpenXmlElement element) =>
            element.Ancestors().Any(ancestor =>
                string.Equals(ancestor.LocalName, "txbxContent", StringComparison.Ordinal) ||
                string.Equals(ancestor.LocalName, "textbox", StringComparison.Ordinal));

        internal sealed class FieldRoot {
            internal FieldRoot(OpenXmlCompositeElement root, WordFieldLocationKind locationKind, string partUri) {
                Root = root;
                LocationKind = locationKind;
                PartUri = partUri;
            }

            internal OpenXmlCompositeElement Root { get; }

            internal WordFieldLocationKind LocationKind { get; }

            internal string PartUri { get; }
        }

        private sealed class FieldInfoCandidate {
            internal FieldInfoCandidate(int sequence, WordFieldInfo info) {
                Sequence = sequence;
                Info = info;
            }

            internal int Sequence { get; }

            private WordFieldInfo Info { get; }

            internal WordFieldInfo InfoWithIndex(int index) {
                Info.Index = index;
                return Info;
            }
        }

        private sealed class ComplexFieldBuilder {
            internal ComplexFieldBuilder(int sequence, int nestingLevel, OpenXmlElement startElement) {
                Sequence = sequence;
                NestingLevel = nestingLevel;
                StartElement = startElement;
            }

            internal int Sequence { get; }

            internal int NestingLevel { get; }

            internal OpenXmlElement StartElement { get; }

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

            internal FieldInfoCandidate ToCandidate(FieldRoot root) {
                string instructionText = string.Concat(InstructionParts);
                ParsedFieldInstruction parsed = ParseInstruction(instructionText);
                var diagnostics = Diagnostics.Concat(parsed.Diagnostics).ToArray();

                var info = new WordFieldInfo(
                    index: 0,
                    representation: WordFieldRepresentation.Complex,
                    locationKind: root.LocationKind,
                    partUri: root.PartUri,
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

                return new FieldInfoCandidate(Sequence, info);
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
