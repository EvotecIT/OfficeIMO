using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Globalization;

namespace OfficeIMO.Word {
    public static partial class WordMailMerge {
        private static readonly Regex MailMergeControlFieldTypePattern = new Regex(
            @"^\s*(?<field>NEXTIF|SKIPIF|NEXT|MERGEREC|MERGESEQ)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));
        private static readonly Regex MergeFieldTypePattern = new Regex(
            @"^\s*MERGEFIELD(?:\s|$)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static void ReplaceMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            ReplaceMergeFields(root, values, removeFields, null);
        }
        private static void ReplaceMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields, List<WordMailMergeFieldResult>? results) {
            foreach (MergeFieldOccurrence occurrence in DiscoverMergeFieldOccurrences(root).OrderBy(item => item.Order)) {
                if (occurrence.SimpleField != null && occurrence.MalformedMessage != null) {
                    ReportMalformedMergeField(
                        results,
                        occurrence.SimpleField.Instruction?.Value ?? string.Empty,
                        occurrence.MalformedMessage);
                } else if (occurrence.SimpleField != null) {
                    ReplaceSimpleMergeField(occurrence.SimpleField, values, removeFields, results);
                } else if (occurrence.MalformedMessage != null) {
                    ReportMalformedMergeField(results, occurrence.ComplexField!.Instruction.ToString(), occurrence.MalformedMessage);
                } else {
                    ReplaceComplexFieldRuns(occurrence.ComplexField!, values, removeFields, results);
                }
            }
        }

        private static void ReplaceSimpleMergeField(SimpleField simpleField, IDictionary<string, string> values, bool removeFields, List<WordMailMergeFieldResult>? results) {
            string instruction = simpleField.Instruction?.Value ?? string.Empty;
            string? name = TryGetMergeFieldName(instruction);
            if (name == null) {
                ReportMalformedMergeField(results, instruction, "A simple MERGEFIELD instruction could not be parsed as a named field.");
                return;
            }
            if (!TryGetMergeValue(values, name, out string? value)) {
                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.MissingValue, null, "Merge field '" + name + "' has no supplied value.");
                return;
            }
            if (!TryFormatMergeValue(instruction, value, out string formattedValue, out string formatMessage)) {
                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.UnsupportedFormatting, null, formatMessage);
                return;
            }

            if (removeFields) {
                var replacement = CreateReplacementRun(formattedValue, EnumerateSimpleFieldOwnedRuns(simpleField).FirstOrDefault());
                simpleField.InsertBeforeSelf(replacement);
                simpleField.Remove();
            } else {
                List<Run> resultRuns = EnumerateSimpleFieldOwnedRuns(simpleField).ToList();
                if (!SetFieldResultText(resultRuns, formattedValue)) {
                    simpleField.Append(CreateReplacementRun(formattedValue, sourceRun: null));
                }
            }
            AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.Merged, formattedValue, "Merge field '" + name + "' was updated.");
        }

        private static IEnumerable<MergeFieldOccurrence> DiscoverMergeFieldOccurrences(OpenXmlElement root) {
            var occurrences = new List<MergeFieldOccurrence>();
            var beginRunOrders = new Dictionary<Run, int>();
            int order = 0;
            foreach (OpenXmlElement element in root.Descendants()) {
                if (element is SimpleField simpleField) {
                    string instruction = simpleField.Instruction?.Value ?? string.Empty;
                    if (!MergeFieldTypePattern.IsMatch(instruction)) {
                        order++;
                        continue;
                    }
                    occurrences.Add(MergeFieldOccurrence.ForSimple(
                        order,
                        simpleField,
                        ContainsNestedField(simpleField)
                            ? "A simple MERGEFIELD contains a nested field and cannot be processed deterministically."
                            : GetMalformedMergeFieldInstructionMessage(instruction)));
                } else if (element is Run run && ContainsFieldMarker(run, FieldCharValues.Begin)) {
                    beginRunOrders[run] = order;
                }
                order++;
            }

            foreach (var paragraph in EnumerateParagraphs(root)) {
                var activeFields = new List<ComplexFieldFrame>();

                foreach (var run in EnumerateParagraphOwnedRuns(paragraph)) {
                    if (activeFields.Count > 0 && run.Ancestors<SimpleField>().Any()) {
                        foreach (ComplexFieldFrame activeField in activeFields) activeField.HasNestedField = true;
                    }

                    FieldChar? firstFieldChar = run.GetFirstChild<FieldChar>();
                    if (activeFields.Count == 0 && firstFieldChar == null) continue;
                    if (activeFields.Count > 0) AddRunToActiveFields(activeFields, run);
                    for (int childIndex = 0; childIndex < run.ChildElements.Count; childIndex++) {
                        OpenXmlElement child = run.ChildElements[childIndex];
                        if (child is FieldCode code) {
                            if (activeFields.Count > 0) activeFields[activeFields.Count - 1].Instruction.Append(code.Text);
                            continue;
                        }
                        if (child is not FieldChar fieldChar) continue;
                        if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin) {
                            foreach (ComplexFieldFrame activeField in activeFields) {
                                activeField.HasNestedField = true;
                            }
                            activeFields.Add(new ComplexFieldFrame(run, fieldChar, beginRunOrders[run]));
                            continue;
                        }
                        if (fieldChar.FieldCharType?.Value != FieldCharValues.End || activeFields.Count == 0) {
                            continue;
                        }

                        ComplexFieldFrame completedField = activeFields[activeFields.Count - 1];
                        activeFields.RemoveAt(activeFields.Count - 1);
                        completedField.EndMarker = fieldChar;
                        string instruction = completedField.Instruction.ToString();
                        if (!MergeFieldTypePattern.IsMatch(instruction)) {
                            continue;
                        }
                        string? malformedMessage = completedField.HasNestedField
                            ? "A complex MERGEFIELD contains a nested field and cannot be processed deterministically."
                            : GetMalformedMergeFieldInstructionMessage(instruction);
                        occurrences.Add(MergeFieldOccurrence.ForComplex(
                            completedField.Order,
                            completedField,
                            malformedMessage));
                    }
                }

                foreach (ComplexFieldFrame activeField in activeFields) {
                    string instruction = activeField.Instruction.ToString();
                    if (!MergeFieldTypePattern.IsMatch(instruction)) {
                        continue;
                    }
                    occurrences.Add(MergeFieldOccurrence.ForComplex(
                        activeField.Order,
                        activeField,
                        "A complex MERGEFIELD is missing its closing field marker or a valid field name."));
                }
            }

            return occurrences;
        }

        private static bool ContainsFieldMarker(Run run, FieldCharValues marker) {
            FieldChar? firstFieldChar = run.GetFirstChild<FieldChar>();
            if (firstFieldChar == null) return false;
            if (firstFieldChar.FieldCharType?.Value == marker) return true;
            for (int index = 0; index < run.ChildElements.Count; index++) {
                if (run.ChildElements[index] is FieldChar fieldChar &&
                    !ReferenceEquals(fieldChar, firstFieldChar) &&
                    fieldChar.FieldCharType?.Value == marker) {
                    return true;
                }
            }
            return false;
        }

        private static void AddRunToActiveFields(IEnumerable<ComplexFieldFrame> activeFields, Run run) {
            foreach (ComplexFieldFrame activeField in activeFields) {
                if (activeField.Runs.Count == 0 ||
                    !ReferenceEquals(activeField.Runs[activeField.Runs.Count - 1], run)) {
                    activeField.Runs.Add(run);
                }
            }
        }

        private static bool ContainsNestedField(SimpleField simpleField) =>
            simpleField.Descendants().Any(element =>
                element is SimpleField or FieldChar or FieldCode);

        private static string? GetMalformedMergeFieldInstructionMessage(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.MergeField || parsed.Instructions.Count != 1) {
                return "A MERGEFIELD instruction must contain exactly one field name.";
            }
            return null;
        }

        private static IEnumerable<Paragraph> EnumerateParagraphs(OpenXmlElement root) {
            if (root is Paragraph paragraph) {
                yield return paragraph;
            }

            foreach (var child in root.Descendants<Paragraph>()) {
                yield return child;
            }
        }

        private static IEnumerable<Run> EnumerateParagraphOwnedRuns(Paragraph paragraph) {
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                foreach (Run run in EnumerateRunsUntilNestedParagraph(child)) {
                    yield return run;
                }
            }
        }

        private static IEnumerable<Run> EnumerateRunsUntilNestedParagraph(OpenXmlElement element) {
            if (element is Paragraph) yield break;
            if (element is Run run) {
                yield return run;
                yield break;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                foreach (Run descendantRun in EnumerateRunsUntilNestedParagraph(child)) {
                    yield return descendantRun;
                }
            }
        }

        private static IEnumerable<Run> EnumerateSimpleFieldOwnedRuns(SimpleField simpleField) {
            foreach (OpenXmlElement child in simpleField.ChildElements) {
                foreach (Run run in EnumerateRunsUntilNestedSimpleField(child)) {
                    yield return run;
                }
            }
        }

        private static IEnumerable<Run> EnumerateRunsUntilNestedSimpleField(OpenXmlElement element) {
            if (element is SimpleField) yield break;
            if (element is Run run) {
                yield return run;
                yield break;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                foreach (Run descendantRun in EnumerateRunsUntilNestedSimpleField(child)) {
                    yield return descendantRun;
                }
            }
        }

        internal static IEnumerable<OpenXmlCompositeElement> EnumerateTemplateRoots(WordDocument document) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            Body? body = mainPart?.Document?.Body;
            if (body != null) {
                yield return body;
            }

            if (mainPart == null) {
                yield break;
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                if (headerPart.Header != null) {
                    yield return headerPart.Header;
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts) {
                if (footerPart.Footer != null) {
                    yield return footerPart.Footer;
                }
            }

            if (mainPart.FootnotesPart?.Footnotes != null) {
                yield return mainPart.FootnotesPart.Footnotes;
            }

            if (mainPart.EndnotesPart?.Endnotes != null) {
                yield return mainPart.EndnotesPart.Endnotes;
            }
        }

        private static void ReplaceComplexFieldRuns(ComplexFieldFrame field, IDictionary<string, string> values, bool removeFields, List<WordMailMergeFieldResult>? results) {
            IReadOnlyList<Run> fieldRuns = field.Runs;
            string instruction = field.Instruction.ToString();
            string? name = TryGetMergeFieldName(instruction);
            if (name == null) {
                ReportMalformedMergeField(results, instruction, "A complex MERGEFIELD instruction could not be parsed as a named field.");
                return;
            }
            if (!TryGetMergeValue(values, name, out string? value)) {
                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.MissingValue, null, "Merge field '" + name + "' has no supplied value.");
                return;
            }
            if (!TryFormatMergeValue(instruction, value, out string formattedValue, out string formatMessage)) {
                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.UnsupportedFormatting, null, formatMessage);
                return;
            }

            if (removeFields) {
                Run? sourceRun = GetComplexFieldResultTextElements(field)
                    .Select(text => text.Parent as Run)
                    .OfType<Run>()
                    .FirstOrDefault()
                    ?? GetComplexFieldResultRuns(field).FirstOrDefault()
                    ?? fieldRuns.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null)
                    ?? fieldRuns.FirstOrDefault();
                if (!ReplaceComplexFieldRangeWithText(field, formattedValue, sourceRun)) {
                    var replacement = CreateReplacementRun(formattedValue, sourceRun);
                    fieldRuns[0].InsertBeforeSelf(replacement);
                    foreach (var run in fieldRuns) {
                        run.Remove();
                    }
                }

                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.Merged, formattedValue, "Merge field '" + name + "' was updated.");
                return;
            }

            if (!SetComplexFieldResultText(field, formattedValue)) {
                if (!InsertComplexFieldResultBeforeEnd(field, formattedValue)) {
                    Run endRun = field.EndMarker?.Parent as Run ?? fieldRuns[fieldRuns.Count - 1];
                    Run? sourceRun = fieldRuns.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null);
                    if (!HasComplexFieldMarker(field, FieldCharValues.Separate)) {
                        endRun.InsertBeforeSelf(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                    }
                    endRun.InsertBeforeSelf(CreateReplacementRun(formattedValue, sourceRun));
                }
            }
            AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.Merged, formattedValue, "Merge field '" + name + "' was updated.");
        }

        private static bool TryFormatMergeValue(string instruction, string value, out string formattedValue, out string message) {
            if (!TryValidateMergeFieldFormattingProfile(instruction, out WordFieldInventory.ParsedFieldInstruction parsed, out message)) {
                formattedValue = string.Empty;
                return false;
            }

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)) {
                if (!decimal.TryParse(value, NumberStyles.Number | NumberStyles.AllowCurrencySymbol, CultureInfo.InvariantCulture, out decimal number)) {
                    formattedValue = string.Empty;
                    message = "Merge field value '" + value + "' is not an invariant number required by numeric picture '" + parsed.NumericPictureSwitch + "'.";
                    return false;
                }
                if (!WordFieldUpdater.TryFormatFormulaValue(number, parsed.NumericPictureSwitch, out formattedValue, out string? diagnostic)) {
                    message = diagnostic ?? "Merge field numeric picture is outside the deterministic formatting profile.";
                    return false;
                }
            } else if (parsed.Switches.Any(fieldSwitch => fieldSwitch.TrimStart().StartsWith(@"\@", StringComparison.Ordinal))) {
                if (!DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out DateTimeOffset dateTime)) {
                    formattedValue = string.Empty;
                    message = "Merge field value '" + value + "' is not an invariant date/time required by the date picture switch.";
                    return false;
                }
                if (!WordFieldUpdater.TryFormatDateTime(dateTime, parsed, out formattedValue, out message)) return false;
            } else {
                formattedValue = value;
            }

            if (!WordFieldUpdater.TryApplyReferenceTextFormat(parsed.FormatSwitches, formattedValue, out string textFormatted, out string? unsupportedFormat)) {
                message = "Merge field text format '" + unsupportedFormat + "' is outside the deterministic formatting profile.";
                return false;
            }
            formattedValue = textFormatted;
            message = string.Empty;
            return true;
        }

        private static bool TryValidateMergeFieldFormattingProfile(
            string instruction,
            out WordFieldInventory.ParsedFieldInstruction parsed,
            out string message) {
            parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.Diagnostics.Count > 0) {
                message = "Merge field formatting is unsupported: " + string.Join(" ", parsed.Diagnostics);
                return false;
            }

            string? unsupportedSwitch = parsed.Switches.FirstOrDefault(fieldSwitch => {
                string trimmed = fieldSwitch.TrimStart();
                return !trimmed.StartsWith(@"\#", StringComparison.Ordinal) &&
                       !trimmed.StartsWith(@"\@", StringComparison.Ordinal) &&
                       !trimmed.StartsWith(@"\*", StringComparison.Ordinal);
            });
            if (unsupportedSwitch != null) {
                message = "Merge field switch '" + unsupportedSwitch.Trim() + "' is outside the deterministic formatting profile.";
                return false;
            }

            bool hasDatePictureSwitch = parsed.Switches.Any(fieldSwitch =>
                fieldSwitch.TrimStart().StartsWith(@"\@", StringComparison.Ordinal));
            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch) && hasDatePictureSwitch) {
                message = "Merge fields cannot combine numeric and date/time picture switches in the deterministic formatting profile.";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)) {
                if (!WordFieldUpdater.TryValidateNumericPictureProfile(parsed.NumericPictureSwitch, out string? diagnostic)) {
                    message = diagnostic ?? "Merge field numeric picture is outside the deterministic formatting profile.";
                    return false;
                }
            } else if (hasDatePictureSwitch) {
                if (!WordFieldUpdater.TryFormatDateTime(new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero), parsed, out _, out message)) return false;
            }

            if (!WordFieldUpdater.TryApplyReferenceTextFormat(parsed.FormatSwitches, string.Empty, out _, out string? unsupportedFormat)) {
                message = "Merge field text format '" + unsupportedFormat + "' is outside the deterministic formatting profile.";
                return false;
            }
            message = string.Empty;
            return true;
        }

        private static void AddMergeResult(List<WordMailMergeFieldResult>? results, string name, string instruction, WordMailMergeFieldStatus status, string? value, string message) {
            results?.Add(new WordMailMergeFieldResult(name, NormalizeFieldInstructionForMessage(instruction), status, value, message));
        }

        private static void ReportMalformedMergeField(List<WordMailMergeFieldResult>? results, string instruction, string message) {
            if (!MergeFieldTypePattern.IsMatch(instruction)) return;
            AddMergeResult(results, string.Empty, instruction, WordMailMergeFieldStatus.MalformedField, null, message);
        }

        private sealed class ComplexFieldFrame {
            internal ComplexFieldFrame(Run beginRun, FieldChar beginMarker, int order) {
                Runs = new List<Run> { beginRun };
                BeginMarker = beginMarker;
                Order = order;
            }

            internal List<Run> Runs { get; }
            internal FieldChar BeginMarker { get; }
            internal FieldChar? EndMarker { get; set; }
            internal StringBuilder Instruction { get; } = new StringBuilder();
            internal int Order { get; }
            internal bool HasNestedField { get; set; }
        }

        private sealed class MergeFieldOccurrence {
            private MergeFieldOccurrence(int order, SimpleField? simpleField, ComplexFieldFrame? complexField, string? malformedMessage) {
                Order = order;
                SimpleField = simpleField;
                ComplexField = complexField;
                MalformedMessage = malformedMessage;
            }

            internal int Order { get; }
            internal SimpleField? SimpleField { get; }
            internal ComplexFieldFrame? ComplexField { get; }
            internal string? MalformedMessage { get; }

            internal static MergeFieldOccurrence ForSimple(int order, SimpleField simpleField, string? malformedMessage = null) =>
                new MergeFieldOccurrence(order, simpleField, null, malformedMessage);

            internal static MergeFieldOccurrence ForComplex(int order, ComplexFieldFrame field, string? malformedMessage = null) =>
                new MergeFieldOccurrence(order, null, field, malformedMessage);
        }

        private static IEnumerable<Run> GetComplexFieldResultRuns(ComplexFieldFrame field) {
            bool insideField = false;
            bool afterSeparator = false;
            foreach (Run run in field.Runs) {
                bool containsResultContent = false;
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (!insideField) {
                        if (ReferenceEquals(child, field.BeginMarker)) insideField = true;
                        continue;
                    }
                    if (ReferenceEquals(child, field.EndMarker)) {
                        if (containsResultContent) yield return run;
                        yield break;
                    }
                    if (child is FieldChar fieldChar &&
                        fieldChar.FieldCharType?.Value == FieldCharValues.Separate) {
                        afterSeparator = true;
                        continue;
                    }
                    if (afterSeparator) containsResultContent = true;
                }
                if (containsResultContent) yield return run;
            }
        }

        private static IEnumerable<Text> GetComplexFieldResultTextElements(
            ComplexFieldFrame field) {
            bool insideField = false;
            bool afterSeparator = false;
            foreach (Run run in field.Runs) {
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (!insideField) {
                        if (ReferenceEquals(child, field.BeginMarker)) insideField = true;
                        continue;
                    }
                    if (ReferenceEquals(child, field.EndMarker)) yield break;
                    if (child is FieldChar fieldChar) {
                        if (fieldChar.FieldCharType?.Value == FieldCharValues.Separate) {
                            afterSeparator = true;
                            continue;
                        }
                    }
                    if (afterSeparator && child is Text text) {
                        yield return text;
                    }
                }
            }
        }

        private static bool SetComplexFieldResultText(
            ComplexFieldFrame field,
            string value) {
            List<Text> markerAwareTextElements = GetComplexFieldResultTextElements(field).ToList();
            return markerAwareTextElements.Count > 0 &&
                SetFieldResultTextElements(markerAwareTextElements, value);
        }

        private static bool InsertComplexFieldResultBeforeEnd(ComplexFieldFrame field, string value) {
            FieldChar? endMarker = field.EndMarker;
            Run? endRun = endMarker?.Parent as Run;
            if (endMarker == null || endRun == null) return false;
            bool inlineBoundary = false;
            foreach (OpenXmlElement child in endRun.ChildElements) {
                if (ReferenceEquals(child, endMarker)) break;
                if (ReferenceEquals(child, field.BeginMarker) || child is FieldCode ||
                    child is FieldChar marker && marker.FieldCharType?.Value == FieldCharValues.Separate) {
                    inlineBoundary = true;
                }
            }
            if (!inlineBoundary) return false;
            if (!HasComplexFieldMarker(field, FieldCharValues.Separate)) {
                endRun.InsertBefore(
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    endMarker);
            }
            endRun.InsertBefore(
                new Text(value) { Space = SpaceProcessingModeValues.Preserve },
                endMarker);
            return true;
        }

        private static bool HasComplexFieldMarker(ComplexFieldFrame field, FieldCharValues marker) {
            bool insideField = false;
            foreach (Run run in field.Runs) {
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (!insideField) {
                        if (ReferenceEquals(child, field.BeginMarker)) insideField = true;
                        continue;
                    }
                    if (ReferenceEquals(child, field.EndMarker)) return marker == FieldCharValues.End;
                    if (child is FieldChar fieldChar && fieldChar.FieldCharType?.Value == marker) return true;
                }
            }
            return false;
        }

        private static bool SetFieldResultText(IEnumerable<Run> runs, string value) {
            List<Run> resultRuns = runs.ToList();
            var textElements = resultRuns
                .SelectMany(run => run.Elements<Text>())
                .ToList();

            if (textElements.Count == 0) {
                if (resultRuns.Count == 0) return false;
                resultRuns[0].Append(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                return true;
            }

            return SetFieldResultTextElements(textElements, value);
        }

        private static bool SetFieldResultTextElements(IReadOnlyList<Text> textElements, string value) {
            textElements[0].Text = value;
            textElements[0].Space = SpaceProcessingModeValues.Preserve;
            for (int i = 1; i < textElements.Count; i++) {
                textElements[i].Text = string.Empty;
            }
            return true;
        }

        private static Run CreateReplacementRun(string value, Run? sourceRun) {
            var run = new Run();
            var properties = sourceRun?.GetFirstChild<RunProperties>();
            if (properties != null) {
                run.Append((RunProperties)properties.CloneNode(true));
            }

            run.Append(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }

        private static bool TryGetMergeValue(IDictionary<string, string> values, string name, out string value) {
            if (values.TryGetValue(name, out value!)) {
                return true;
            }

            foreach (KeyValuePair<string, string> entry in values) {
                if (string.Equals(entry.Key, name, StringComparison.OrdinalIgnoreCase)) {
                    value = entry.Value;
                    return true;
                }
            }

            value = string.Empty;
            return false;
        }

        private static string? GetMergeFieldName(string? fieldInstruction) {
            if (string.IsNullOrWhiteSpace(fieldInstruction)) {
                return null;
            }

            var parser = new WordFieldParser(fieldInstruction!);
            if (parser.WordFieldType != WordFieldType.MergeField || parser.Instructions.Count != 1) {
                return null;
            }

            return parser.Instructions[0].Trim().Trim('"');
        }

        private static IEnumerable<string> EnumerateMergeFieldNames(OpenXmlElement root) {
            foreach (var simpleField in root.Descendants<SimpleField>()) {
                string? name = TryGetMergeFieldName(simpleField.Instruction?.Value);
                if (!string.IsNullOrWhiteSpace(name)) {
                    yield return name!;
                }
            }

            foreach (string instruction in EnumerateComplexFieldInstructions(root)) {
                string? name = TryGetMergeFieldName(instruction);
                if (!string.IsNullOrWhiteSpace(name)) {
                    yield return name!;
                }
            }
        }

        private static string? TryGetMergeFieldName(string? fieldInstruction) {
            try {
                return GetMergeFieldName(fieldInstruction);
            } catch (NotImplementedException) {
                return null;
            }
        }

        private static IEnumerable<WordMailMergeTemplateIssue> EnumerateMalformedMergeFieldIssues(OpenXmlElement root) {
            foreach (MergeFieldOccurrence occurrence in DiscoverMergeFieldOccurrences(root)) {
                if (occurrence.MalformedMessage == null) continue;
                string instruction = occurrence.SimpleField?.Instruction?.Value
                    ?? occurrence.ComplexField!.Instruction.ToString();
                if (!MergeFieldTypePattern.IsMatch(instruction)) continue;

                string name = TryGetMergeFieldName(instruction) ?? string.Empty;
                yield return new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.MalformedMergeField,
                    name,
                    occurrence.MalformedMessage);
            }
        }

        private static IEnumerable<WordMailMergeTemplateIssue> EnumerateUnsupportedMergeFieldFormattingIssues(OpenXmlElement root) {
            foreach (MergeFieldOccurrence occurrence in DiscoverMergeFieldOccurrences(root)) {
                if (occurrence.MalformedMessage != null) continue;
                string instruction = occurrence.SimpleField?.Instruction?.Value
                    ?? occurrence.ComplexField!.Instruction.ToString();
                if (!MergeFieldTypePattern.IsMatch(instruction)) continue;

                string? name = TryGetMergeFieldName(instruction);
                if (string.IsNullOrWhiteSpace(name) ||
                    TryValidateMergeFieldFormattingProfile(instruction, out _, out string message)) {
                    continue;
                }

                yield return new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.UnsupportedMergeFieldFormatting,
                    name!,
                    message);
            }
        }

        private static IEnumerable<WordMailMergeTemplateIssue> EnumerateUnsupportedMailMergeControlFieldIssues(OpenXmlElement root) {
            foreach (string instruction in EnumerateFieldInstructions(root)) {
                if (!TryGetUnsupportedMailMergeControlField(instruction, out string? fieldName)) {
                    continue;
                }

                yield return new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.UnsupportedMailMergeControlField,
                    fieldName!,
                    $"{fieldName} field '{NormalizeFieldInstructionForMessage(instruction)}' is a Word-native mail-merge record-control field and is not executed by OfficeIMO mail merge.");
            }
        }

        private static IEnumerable<string> EnumerateFieldInstructions(OpenXmlElement root) {
            foreach (var simpleField in root.Descendants<SimpleField>()) {
                string? instruction = simpleField.Instruction?.Value;
                if (!string.IsNullOrWhiteSpace(instruction)) {
                    yield return instruction!;
                }
            }

            foreach (string instruction in EnumerateComplexFieldInstructions(root)) {
                if (!string.IsNullOrWhiteSpace(instruction)) {
                    yield return instruction;
                }
            }
        }

        private static IEnumerable<string> EnumerateComplexFieldInstructions(OpenXmlElement root) {
            foreach (Paragraph paragraph in EnumerateParagraphs(root)) {
                var activeFields = new List<StringBuilder>();
                foreach (Run run in EnumerateParagraphOwnedRuns(paragraph)) {
                    if (activeFields.Count == 0 && run.GetFirstChild<FieldChar>() == null) continue;
                    for (int childIndex = 0; childIndex < run.ChildElements.Count; childIndex++) {
                        OpenXmlElement child = run.ChildElements[childIndex];
                        if (child is FieldChar fieldChar) {
                            if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin) {
                                activeFields.Add(new StringBuilder());
                            } else if (fieldChar.FieldCharType?.Value == FieldCharValues.End &&
                                       activeFields.Count > 0) {
                                StringBuilder completed = activeFields[activeFields.Count - 1];
                                activeFields.RemoveAt(activeFields.Count - 1);
                                yield return completed.ToString();
                            }
                        } else if (child is FieldCode code && activeFields.Count > 0) {
                            activeFields[activeFields.Count - 1].Append(code.Text);
                        }
                    }
                }
            }
        }

        private static bool TryGetUnsupportedMailMergeControlField(string? instruction, out string? fieldName) {
            fieldName = null;
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            Match match = MailMergeControlFieldTypePattern.Match(instruction!);
            if (!match.Success) {
                return false;
            }

            fieldName = match.Groups["field"].Value.ToUpperInvariant();
            return true;
        }

        private static string NormalizeFieldInstructionForMessage(string instruction) {
            return Regex.Replace(instruction.Trim(), @"\s+", " ");
        }

    }
}
