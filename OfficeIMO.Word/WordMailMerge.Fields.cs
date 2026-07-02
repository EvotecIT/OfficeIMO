using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeIMO.Word {
    public static partial class WordMailMerge {
        private static readonly Regex MailMergeControlFieldTypePattern = new Regex(
            @"^\s*(?<field>NEXTIF|SKIPIF|NEXT|MERGEREC|MERGESEQ)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static void ReplaceMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            ReplaceSimpleMergeFields(root, values, removeFields);
            ReplaceComplexMergeFields(root, values, removeFields);
        }
        private static void ReplaceSimpleMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            foreach (var simpleField in root.Descendants<SimpleField>().ToList()) {
                string? name = GetMergeFieldName(simpleField.Instruction?.Value);
                if (name == null || !TryGetMergeValue(values, name, out string? value)) {
                    continue;
                }

                if (removeFields) {
                    var replacement = CreateReplacementRun(value, simpleField.Elements<Run>().FirstOrDefault());
                    simpleField.InsertBeforeSelf(replacement);
                    simpleField.Remove();
                } else {
                    SetFieldResultText(simpleField.Elements<Run>(), value);
                }
            }
        }

        private static void ReplaceComplexMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            foreach (var paragraph in EnumerateParagraphs(root)) {
                List<Run>? fieldRuns = null;

                foreach (var run in paragraph.Elements<Run>().ToList()) {
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

                    ReplaceComplexFieldRuns(fieldRuns, values, removeFields);
                    fieldRuns = null;
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

        private static void ReplaceComplexFieldRuns(IReadOnlyList<Run> fieldRuns, IDictionary<string, string> values, bool removeFields) {
            string instruction = string.Concat(fieldRuns
                .SelectMany(run => run.Elements<FieldCode>())
                .Select(code => code.Text));
            string? name = GetMergeFieldName(instruction);
            if (name == null || !TryGetMergeValue(values, name, out string? value)) {
                return;
            }

            if (removeFields) {
                Run? sourceRun = GetComplexFieldResultRuns(fieldRuns).FirstOrDefault()
                    ?? fieldRuns.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null)
                    ?? fieldRuns.FirstOrDefault();
                var replacement = CreateReplacementRun(value, sourceRun);
                fieldRuns[0].InsertBeforeSelf(replacement);
                foreach (var run in fieldRuns) {
                    run.Remove();
                }

                return;
            }

            var resultRuns = GetComplexFieldResultRuns(fieldRuns).ToList();
            SetFieldResultText(resultRuns, value);
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

        private static void SetFieldResultText(IEnumerable<Run> runs, string value) {
            var textElements = runs
                .SelectMany(run => run.Elements<Text>())
                .ToList();

            if (textElements.Count == 0) {
                return;
            }

            textElements[0].Text = value;
            textElements[0].Space = SpaceProcessingModeValues.Preserve;
            for (int i = 1; i < textElements.Count; i++) {
                textElements[i].Text = string.Empty;
            }
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
