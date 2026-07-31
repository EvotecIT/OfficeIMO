using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
            ReplaceSimpleMergeFields(root, values, removeFields, results);
            ReplaceComplexMergeFields(root, values, removeFields, results);
        }
        private static void ReplaceSimpleMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields, List<WordMailMergeFieldResult>? results) {
            foreach (var simpleField in root.Descendants<SimpleField>().ToList()) {
                string instruction = simpleField.Instruction?.Value ?? string.Empty;
                string? name = TryGetMergeFieldName(instruction);
                if (name == null) {
                    ReportMalformedMergeField(results, instruction, "A simple MERGEFIELD instruction could not be parsed as a named field.");
                    continue;
                }
                if (!TryGetMergeValue(values, name, out string? value)) {
                    AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.MissingValue, null, "Merge field '" + name + "' has no supplied value.");
                    continue;
                }
                if (!TryFormatMergeValue(instruction, value, out string formattedValue, out string formatMessage)) {
                    AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.UnsupportedFormatting, null, formatMessage);
                    continue;
                }

                if (removeFields) {
                    var replacement = CreateReplacementRun(formattedValue, simpleField.Elements<Run>().FirstOrDefault());
                    simpleField.InsertBeforeSelf(replacement);
                    simpleField.Remove();
                } else {
                    List<Run> resultRuns = simpleField.Elements<Run>().ToList();
                    if (!SetFieldResultText(resultRuns, formattedValue)) {
                        simpleField.Append(CreateReplacementRun(formattedValue, sourceRun: null));
                    }
                }
                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.Merged, formattedValue, "Merge field '" + name + "' was updated.");
            }
        }

        private static void ReplaceComplexMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields, List<WordMailMergeFieldResult>? results) {
            foreach (var paragraph in EnumerateParagraphs(root)) {
                var activeFields = new List<ComplexFieldFrame>();

                foreach (var run in paragraph.Elements<Run>().ToList()) {
                    var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                    if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                        foreach (ComplexFieldFrame activeField in activeFields) {
                            activeField.Runs.Add(run);
                            activeField.HasNestedField = true;
                        }
                        activeFields.Add(new ComplexFieldFrame(run));
                        continue;
                    }

                    if (activeFields.Count == 0) {
                        continue;
                    }

                    foreach (ComplexFieldFrame activeField in activeFields) activeField.Runs.Add(run);
                    if (fieldChar?.FieldCharType?.Value != FieldCharValues.End) {
                        continue;
                    }

                    ComplexFieldFrame completedField = activeFields[activeFields.Count - 1];
                    activeFields.RemoveAt(activeFields.Count - 1);
                    string instruction = ReadComplexFieldInstruction(completedField.Runs);
                    if (completedField.HasNestedField && MergeFieldTypePattern.IsMatch(instruction)) {
                        ReportMalformedMergeField(results, instruction, "A complex MERGEFIELD contains a nested field and cannot be processed deterministically.");
                    } else {
                        ReplaceComplexFieldRuns(completedField.Runs, values, removeFields, results);
                    }
                }

                foreach (ComplexFieldFrame activeField in activeFields) {
                    string instruction = ReadComplexFieldInstruction(activeField.Runs);
                    ReportMalformedMergeField(results, instruction, "A complex MERGEFIELD is missing its closing field marker or a valid field name.");
                }
            }
        }

        private static IEnumerable<Paragraph> EnumerateParagraphs(OpenXmlElement root) {
            if (root is Paragraph paragraph) {
                yield return paragraph;
            }

            foreach (var child in root.Descendants<Paragraph>()) {
                yield return child;
            }
        }

        private static IEnumerable<OpenXmlCompositeElement> EnumerateTemplateRoots(WordDocument document) {
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

        private static void ReplaceComplexFieldRuns(IReadOnlyList<Run> fieldRuns, IDictionary<string, string> values, bool removeFields, List<WordMailMergeFieldResult>? results) {
            string instruction = ReadComplexFieldInstruction(fieldRuns);
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
                Run? sourceRun = GetComplexFieldResultRuns(fieldRuns).FirstOrDefault()
                    ?? fieldRuns.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null)
                    ?? fieldRuns.FirstOrDefault();
                var replacement = CreateReplacementRun(formattedValue, sourceRun);
                fieldRuns[0].InsertBeforeSelf(replacement);
                foreach (var run in fieldRuns) {
                    run.Remove();
                }

                AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.Merged, formattedValue, "Merge field '" + name + "' was updated.");
                return;
            }

            var resultRuns = GetComplexFieldResultRuns(fieldRuns).ToList();
            if (!SetFieldResultText(resultRuns, formattedValue)) {
                Run endRun = fieldRuns[fieldRuns.Count - 1];
                Run? sourceRun = fieldRuns.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null);
                endRun.InsertBeforeSelf(CreateReplacementRun(formattedValue, sourceRun));
            }
            AddMergeResult(results, name, instruction, WordMailMergeFieldStatus.Merged, formattedValue, "Merge field '" + name + "' was updated.");
        }

        private static bool TryFormatMergeValue(string instruction, string value, out string formattedValue, out string message) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.Diagnostics.Count > 0) {
                formattedValue = string.Empty;
                message = "Merge field formatting is unsupported: " + string.Join(" ", parsed.Diagnostics);
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
                if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind, out DateTime dateTime)) {
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

        private static void AddMergeResult(List<WordMailMergeFieldResult>? results, string name, string instruction, WordMailMergeFieldStatus status, string? value, string message) {
            results?.Add(new WordMailMergeFieldResult(name, NormalizeFieldInstructionForMessage(instruction), status, value, message));
        }

        private static string ReadComplexFieldInstruction(IEnumerable<Run> fieldRuns) => string.Concat(fieldRuns
            .SelectMany(run => run.Elements<FieldCode>())
            .Select(code => code.Text));

        private static void ReportMalformedMergeField(List<WordMailMergeFieldResult>? results, string instruction, string message) {
            if (!MergeFieldTypePattern.IsMatch(instruction)) return;
            AddMergeResult(results, string.Empty, instruction, WordMailMergeFieldStatus.MalformedField, null, message);
        }

        private sealed class ComplexFieldFrame {
            internal ComplexFieldFrame(Run beginRun) {
                Runs = new List<Run> { beginRun };
            }

            internal List<Run> Runs { get; }
            internal bool HasNestedField { get; set; }
        }

        private static IEnumerable<Run> GetComplexFieldResultRuns(IReadOnlyList<Run> fieldRuns) {
            bool afterSeparator = false;

            foreach (var run in fieldRuns) {
                var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                if (fieldChar?.FieldCharType?.Value == FieldCharValues.Separate) {
                    afterSeparator = true;
                    continue;
                }

                if (fieldChar?.FieldCharType?.Value == FieldCharValues.End) {
                    yield break;
                }

                if (afterSeparator) {
                    yield return run;
                }
            }
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
            if (parser.WordFieldType != WordFieldType.MergeField || parser.Instructions.Count == 0) {
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

            foreach (var paragraph in EnumerateParagraphs(root)) {
                List<Run>? fieldRuns = null;
                foreach (var run in paragraph.Elements<Run>()) {
                    var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                    if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                        fieldRuns = new List<Run> { run };
                        continue;
                    }

                    if (fieldRuns == null) {
                        continue;
                    }

                    fieldRuns.Add(run);
                    if (fieldChar?.FieldCharType?.Value != FieldCharValues.End) {
                        continue;
                    }

                    string instruction = string.Concat(fieldRuns
                        .SelectMany(item => item.Elements<FieldCode>())
                        .Select(code => code.Text));
                    string? name = TryGetMergeFieldName(instruction);
                    if (!string.IsNullOrWhiteSpace(name)) {
                        yield return name!;
                    }

                    fieldRuns = null;
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

            foreach (var paragraph in EnumerateParagraphs(root)) {
                List<Run>? fieldRuns = null;
                foreach (var run in paragraph.Elements<Run>()) {
                    var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                    if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                        fieldRuns = new List<Run> { run };
                        continue;
                    }

                    if (fieldRuns == null) {
                        continue;
                    }

                    fieldRuns.Add(run);
                    if (fieldChar?.FieldCharType?.Value != FieldCharValues.End) {
                        continue;
                    }

                    string instruction = string.Concat(fieldRuns
                        .SelectMany(item => item.Elements<FieldCode>())
                        .Select(code => code.Text));
                    if (!string.IsNullOrWhiteSpace(instruction)) {
                        yield return instruction;
                    }

                    fieldRuns = null;
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
