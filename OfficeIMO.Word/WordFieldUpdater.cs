using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private const string DefaultDateTimeFormat = "yyyy-MM-dd HH:mm:ss";

        internal static WordFieldUpdateReport Update(WordDocument document, WordFieldUpdateOptions options) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");

            int totalPages = EstimateTotalPages(document);
            DateTime updateDateTime = options.CurrentDateTime ?? DateTime.Now;
            var state = new FieldEvaluationState();
            var results = new List<WordFieldUpdateResult>();
            var replacedContainingFields = new List<MutableFieldCandidate>();
            List<MutableFieldCandidate> candidates = EnumerateFields(mainPart).OrderBy(field => field.Sequence).ToList();

            for (int index = 0; index < candidates.Count; index++) {
                candidates[index].Index = index;
            }

            foreach (MutableFieldCandidate candidate in SortFieldsForEvaluation(candidates)) {
                if (IsNestedInsideReplacedField(candidate, replacedContainingFields)) {
                    results.Add(candidate.ToResult(
                        WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType,
                        WordFieldUpdateStatus.Skipped,
                        null,
                        "Nested field was left unchanged because its containing field result was replaced."));
                    continue;
                }

                WordFieldUpdateResult result = UpdateField(document, candidate, totalPages, state, updateDateTime);
                results.Add(result);
                if (result.Status == WordFieldUpdateStatus.Updated &&
                    candidate.Representation == WordFieldRepresentation.Complex &&
                    candidate.EndRun != null) {
                    replacedContainingFields.Add(candidate);
                }
            }

            document.TableOfContent?.Update();

            return new WordFieldUpdateReport(results.OrderBy(result => result.Index).ToArray());
        }

        private static WordFieldUpdateResult UpdateField(WordDocument document, MutableFieldCandidate candidate, int totalPages, FieldEvaluationState state, DateTime updateDateTime) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(candidate.InstructionText);

            if (parsed.Diagnostics.Count > 0 || parsed.FieldType == null) {
                return candidate.ToResult(
                    parsed.FieldType,
                    parsed.FieldType == null ? WordFieldUpdateStatus.ParseError : WordFieldUpdateStatus.Unsupported,
                    null,
                    parsed.Diagnostics.Count == 0 ? "Field instruction could not be parsed." : string.Join(" ", parsed.Diagnostics));
            }

            if (candidate.IsLocked) {
                return candidate.ToResult(parsed.FieldType, WordFieldUpdateStatus.Skipped, null, "Field is locked and was left unchanged.");
            }

            if (!TryEvaluate(document, candidate, parsed, totalPages, state, updateDateTime, out string? value, out WordFieldUpdateStatus status, out string message)) {
                return candidate.ToResult(parsed.FieldType, status, null, message);
            }

            SetResultText(candidate, value!);
            return candidate.ToResult(parsed.FieldType, WordFieldUpdateStatus.Updated, value, message);
        }

        private static bool TryEvaluate(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            int totalPages,
            FieldEvaluationState state,
            DateTime updateDateTime,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = $"Field type {parsed.FieldType} is not evaluated by OfficeIMO.";

            switch (parsed.FieldType) {
                case WordFieldType.Author:
                    return TrySetTextValue(document.BuiltinDocumentProperties.Creator, parsed, "Updated from built-in document property Creator.", out value, out status, out message);
                case WordFieldType.Title:
                    return TrySetTextValue(document.BuiltinDocumentProperties.Title, parsed, "Updated from built-in document property Title.", out value, out status, out message);
                case WordFieldType.Subject:
                    return TrySetTextValue(document.BuiltinDocumentProperties.Subject, parsed, "Updated from built-in document property Subject.", out value, out status, out message);
                case WordFieldType.Keywords:
                    return TrySetTextValue(document.BuiltinDocumentProperties.Keywords, parsed, "Updated from built-in document property Keywords.", out value, out status, out message);
                case WordFieldType.Comments:
                    return TrySetTextValue(document.BuiltinDocumentProperties.Description, parsed, "Updated from built-in document property Description.", out value, out status, out message);
                case WordFieldType.LastSavedBy:
                    return TrySetTextValue(document.BuiltinDocumentProperties.LastModifiedBy, parsed, "Updated from built-in document property LastModifiedBy.", out value, out status, out message);
                case WordFieldType.CreateDate:
                    return TrySetDate(document.BuiltinDocumentProperties.Created, parsed, "Updated from built-in document property Created.", out value, out status, out message);
                case WordFieldType.Date:
                    return TrySetDate(updateDateTime, parsed, "Updated from field update date/time.", out value, out status, out message);
                case WordFieldType.PrintDate:
                    return TrySetDate(document.BuiltinDocumentProperties.LastPrinted, parsed, "Updated from built-in document property LastPrinted.", out value, out status, out message);
                case WordFieldType.SaveDate:
                    return TrySetDate(document.BuiltinDocumentProperties.Modified, parsed, "Updated from built-in document property Modified.", out value, out status, out message);
                case WordFieldType.Time:
                    return TrySetDate(updateDateTime, parsed, "Updated from field update date/time.", out value, out status, out message);
                case WordFieldType.FileName:
                    return TryEvaluateFileName(document, parsed, out value, out status, out message);
                case WordFieldType.FileSize:
                    return TryEvaluateFileSize(document, parsed, out value, out status, out message);
                case WordFieldType.Info:
                    return TryEvaluateInfo(document, parsed, out value, out status, out message);
                case WordFieldType.DocProperty:
                    return TryEvaluateDocumentProperty(document, parsed, out value, out status, out message);
                case WordFieldType.DocVariable:
                    return TryEvaluateDocumentVariable(document, parsed, out value, out status, out message);
                case WordFieldType.Quote:
                    return TryEvaluateQuote(parsed, out value, out status, out message);
                case WordFieldType.Ref:
                    return TryEvaluateRef(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.PageRef:
                    return TryEvaluatePageRef(document, parsed, out value, out status, out message);
                case WordFieldType.Formula:
                    return TryEvaluateFormula(candidate, parsed, out value, out status, out message);
                case WordFieldType.Seq:
                    return TryEvaluateSequence(document, candidate, parsed, state, out value, out status, out message);
                case WordFieldType.Page:
                    return TryEvaluatePage(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.NumChars:
                    return TryFormatNumericField(document.Statistics.Characters, parsed.NumericPictureSwitch, "Updated from OfficeIMO document statistics Characters.", out value, out status, out message);
                case WordFieldType.NumWords:
                    return TryFormatNumericField(document.Statistics.Words, parsed.NumericPictureSwitch, "Updated from OfficeIMO document statistics Words.", out value, out status, out message);
                case WordFieldType.NumPages:
                    return TryFormatNumericField(totalPages, parsed.NumericPictureSwitch, "Updated from OfficeIMO page-break count.", out value, out status, out message);
                case WordFieldType.RevNum:
                    return TrySetTextValue(document.BuiltinDocumentProperties.Revision, parsed, "Updated from built-in document property Revision.", out value, out status, out message);
                case WordFieldType.Section:
                    return TryEvaluateSectionNumber(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.SectionPages:
                    return TryEvaluateSectionPages(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.TOC:
                    status = WordFieldUpdateStatus.Skipped;
                    message = "Table of contents refresh was queued for Word to update on open; call WordTableOfContent.RefreshEntries() to generate deterministic OfficeIMO entries.";
                    return false;
                default:
                    return false;
            }
        }

        private static bool TryEvaluateInfo(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "INFO field is missing a built-in property name.";

            string? propertyName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(propertyName)) {
                return false;
            }

            propertyName = TrimQuotes(propertyName);
            if (TryEvaluateBuiltInDocumentProperty(document, propertyName, parsed, out value, out status, out message)) {
                return true;
            }

            if (status == WordFieldUpdateStatus.Unsupported) {
                return false;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = $"Built-in document property {propertyName} was not found or is empty.";
            return false;
        }

        private static bool TryEvaluateDocumentProperty(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "DOCPROPERTY field is missing a property name.";

            string? propertyName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(propertyName)) {
                return false;
            }

            propertyName = TrimQuotes(propertyName);

            if (TryEvaluateBuiltInDocumentProperty(document, propertyName, parsed, out value, out status, out message)) {
                return true;
            }

            if (status == WordFieldUpdateStatus.Unsupported) {
                return false;
            }

            KeyValuePair<string, WordCustomProperty> customProperty = document.CustomDocumentProperties
                .FirstOrDefault(pair => string.Equals(pair.Key, propertyName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(customProperty.Key)) {
                return TrySetTextValue(FormatValue(customProperty.Value.Value), parsed, $"Updated from custom document property {customProperty.Key}.", out value, out status, out message);
            }

            status = WordFieldUpdateStatus.Skipped;
            message = $"Document property {propertyName} was not found.";
            return false;
        }

        private static bool TryEvaluateDocumentVariable(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "DOCVARIABLE field is missing a variable name.";

            string? variableName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(variableName)) {
                return false;
            }

            variableName = TrimQuotes(variableName);
            if (string.IsNullOrWhiteSpace(variableName)) {
                return false;
            }

            KeyValuePair<string, string> variable = document.DocumentVariables
                .FirstOrDefault(pair => string.Equals(pair.Key, variableName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(variable.Key)) {
                value = variable.Value;
                status = WordFieldUpdateStatus.Updated;
                message = $"Updated from document variable {variable.Key}.";
                return true;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = $"Document variable {variableName} was not found.";
            return false;
        }

        private static bool TryEvaluateQuote(
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "QUOTE fields need exactly one quoted literal instruction to be evaluated by OfficeIMO.";

            if (parsed.Instructions.Count != 1) {
                return false;
            }

            if (parsed.Switches.Count > 0) {
                message = "QUOTE fields with general switches are not evaluated by OfficeIMO.";
                return false;
            }

            string literal = parsed.Instructions[0].Trim();
            if (literal.Length < 2 || literal[0] != '"' || literal[literal.Length - 1] != '"') {
                message = "QUOTE fields need a quoted literal instruction to be evaluated by OfficeIMO.";
                return false;
            }

            string literalValue = TrimQuotes(literal);
            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)) {
                if (GetLastMeaningfulFormat(parsed.FormatSwitches) != null) {
                    message = @"QUOTE fields cannot combine \# numeric picture and \* format switches for deterministic literal refresh.";
                    return false;
                }

                if (!decimal.TryParse(literalValue, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal numericLiteral)) {
                    message = "QUOTE field numeric picture switch requires a numeric literal.";
                    return false;
                }

                if (!TryFormatFormulaValue(numericLiteral, parsed.NumericPictureSwitch, out string numericPictureValue, out string? numericPictureDiagnostic)) {
                    message = numericPictureDiagnostic == null
                        ? "QUOTE field numeric picture switch could not be applied."
                        : numericPictureDiagnostic.Replace("Formula numeric picture", "Field numeric picture");
                    return false;
                }

                value = numericPictureValue;
                status = WordFieldUpdateStatus.Updated;
                message = "Updated from QUOTE numeric literal instruction with numeric picture formatting.";
                return true;
            }

            if (!TryApplyQuoteFormat(parsed.FormatSwitches, literalValue, out string formattedValue, out string? unsupportedFormat, out bool usedNumericFormat)) {
                message = $"QUOTE field format switch {unsupportedFormat} is not supported for literal refresh.";
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = parsed.FormatSwitches.Any(fieldFormat => fieldFormat != WordFieldFormat.Mergeformat && fieldFormat != WordFieldFormat.CharFormat)
                ? usedNumericFormat
                    ? "Updated from QUOTE numeric literal instruction with deterministic number formatting."
                    : "Updated from QUOTE literal instruction with deterministic text formatting."
                : "Updated from QUOTE literal instruction.";
            return true;
        }

        private static bool TryApplyQuoteFormat(
            IReadOnlyList<WordFieldFormat> formatSwitches,
            string source,
            out string value,
            out string? unsupportedFormat,
            out bool usedNumericFormat) {
            usedNumericFormat = false;
            WordFieldFormat? format = GetLastMeaningfulFormat(formatSwitches);
            switch (format) {
                case WordFieldFormat.Roman:
                case WordFieldFormat.roman:
                case WordFieldFormat.Arabic:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.Alphabetical:
                case WordFieldFormat.ALPHABETICAL:
                case WordFieldFormat.Hex:
                case WordFieldFormat.CardText:
                case WordFieldFormat.OrdText:
                case WordFieldFormat.DollarText:
                    if (!int.TryParse(source, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numericValue)) {
                        value = source;
                        unsupportedFormat = format.Value.ToString();
                        return false;
                    }

                    if (RequiresNonNegativeNumber(format.Value) && numericValue < 0) {
                        value = source;
                        unsupportedFormat = format.Value.ToString();
                        return false;
                    }

                    value = FormatSequenceValue(numericValue, new[] { format.Value });
                    unsupportedFormat = null;
                    usedNumericFormat = true;
                    return true;
                default:
                    usedNumericFormat = false;
                    return TryApplyReferenceTextFormat(formatSwitches, source, out value, out unsupportedFormat);
            }
        }

        private static bool RequiresNonNegativeNumber(WordFieldFormat format) {
            return format == WordFieldFormat.Hex ||
                format == WordFieldFormat.CardText ||
                format == WordFieldFormat.OrdText ||
                format == WordFieldFormat.DollarText;
        }

        private static IReadOnlyList<MutableFieldCandidate> SortFieldsForEvaluation(IReadOnlyList<MutableFieldCandidate> candidates) {
            var instructionChildren = candidates
                .Where(candidate => candidate.InstructionParentSequence.HasValue)
                .GroupBy(candidate => candidate.InstructionParentSequence!.Value)
                .ToDictionary(group => group.Key, group => group.OrderBy(candidate => candidate.Sequence).ToArray());

            var ordered = new List<MutableFieldCandidate>(candidates.Count);
            var visited = new HashSet<int>();

            foreach (MutableFieldCandidate candidate in candidates.OrderBy(candidate => candidate.Sequence)) {
                Visit(candidate);
            }

            return ordered;

            void Visit(MutableFieldCandidate candidate) {
                if (!visited.Add(candidate.Sequence)) {
                    return;
                }

                if (instructionChildren.TryGetValue(candidate.Sequence, out MutableFieldCandidate[]? children)) {
                    foreach (MutableFieldCandidate child in children) {
                        Visit(child);
                    }
                }

                ordered.Add(candidate);
            }
        }

        private static bool TryEvaluateBuiltInDocumentProperty(
            WordDocument document,
            string propertyName,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;
            message = string.Empty;

            switch (propertyName.Replace(" ", string.Empty).ToUpperInvariant()) {
                case "AUTHOR":
                case "CREATOR":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Creator, parsed, "Updated from built-in document property Creator.", out value, out status, out message);
                case "TITLE":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Title, parsed, "Updated from built-in document property Title.", out value, out status, out message);
                case "SUBJECT":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Subject, parsed, "Updated from built-in document property Subject.", out value, out status, out message);
                case "CATEGORY":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Category, parsed, "Updated from built-in document property Category.", out value, out status, out message);
                case "KEYWORDS":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Keywords, parsed, "Updated from built-in document property Keywords.", out value, out status, out message);
                case "COMMENTS":
                case "DESCRIPTION":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Description, parsed, "Updated from built-in document property Description.", out value, out status, out message);
                case "LASTSAVEDBY":
                case "LASTMODIFIEDBY":
                    return TrySetTextValue(document.BuiltinDocumentProperties.LastModifiedBy, parsed, "Updated from built-in document property LastModifiedBy.", out value, out status, out message);
                case "LASTPRINTED":
                case "PRINTDATE":
                    return TrySetTextValue(FormatValue(document.BuiltinDocumentProperties.LastPrinted), parsed, "Updated from built-in document property LastPrinted.", out value, out status, out message);
                case "CREATED":
                case "CREATEDATE":
                    return TrySetTextValue(FormatValue(document.BuiltinDocumentProperties.Created), parsed, "Updated from built-in document property Created.", out value, out status, out message);
                case "MODIFIED":
                case "SAVEDATE":
                    return TrySetTextValue(FormatValue(document.BuiltinDocumentProperties.Modified), parsed, "Updated from built-in document property Modified.", out value, out status, out message);
                case "REVISION":
                case "REVNUM":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Revision, parsed, "Updated from built-in document property Revision.", out value, out status, out message);
                case "VERSION":
                    return TrySetTextValue(document.BuiltinDocumentProperties.Version, parsed, "Updated from built-in document property Version.", out value, out status, out message);
                default:
                    return false;
            }
        }

        private static bool TryEvaluateFileName(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (string.IsNullOrWhiteSpace(document.FilePath)) {
                message = "FILENAME cannot be evaluated because the document has no backing file path.";
                return false;
            }

            bool includePath = parsed.Switches.Any(fieldSwitch => string.Equals(fieldSwitch, "\\p", StringComparison.OrdinalIgnoreCase));
            value = includePath ? document.FilePath : Path.GetFileName(document.FilePath);
            status = WordFieldUpdateStatus.Updated;
            message = includePath ? "Updated from document file path." : "Updated from document file name.";
            return true;
        }

        private static bool TryEvaluateFileSize(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (string.IsNullOrWhiteSpace(document.FilePath)) {
                message = "FILESIZE cannot be evaluated because the document has no backing file path.";
                return false;
            }

            if (!File.Exists(document.FilePath)) {
                message = "FILESIZE cannot be evaluated because the backing file does not exist.";
                return false;
            }

            if (!TryGetFileSizeUnit(parsed.Switches, out FileSizeUnit unit, out string? unsupportedSwitch)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"FILESIZE switch {unsupportedSwitch} is not supported for deterministic refresh.";
                return false;
            }

            long bytes = new FileInfo(document.FilePath).Length;
            decimal numericValue = unit switch {
                FileSizeUnit.Kilobytes => Math.Round(bytes / 1000m, 0, MidpointRounding.AwayFromZero),
                FileSizeUnit.Megabytes => Math.Round(bytes / 1000000m, 0, MidpointRounding.AwayFromZero),
                _ => bytes
            };

            string unitMessage = unit switch {
                FileSizeUnit.Kilobytes => "rounded decimal kilobytes",
                FileSizeUnit.Megabytes => "rounded decimal megabytes",
                _ => "bytes"
            };

            return TryFormatNumericField(numericValue, parsed.NumericPictureSwitch, $"Updated from backing DOCX package file size in {unitMessage}.", out value, out status, out message);
        }

        private static bool TryGetFileSizeUnit(IReadOnlyList<string> switches, out FileSizeUnit unit, out string? unsupportedSwitch) {
            unit = FileSizeUnit.Bytes;
            unsupportedSwitch = null;

            foreach (string fieldSwitch in switches) {
                string normalizedSwitch = fieldSwitch.Trim();
                if (string.Equals(normalizedSwitch, "\\k", StringComparison.OrdinalIgnoreCase)) {
                    unit = FileSizeUnit.Kilobytes;
                    continue;
                }

                if (string.Equals(normalizedSwitch, "\\m", StringComparison.OrdinalIgnoreCase)) {
                    unit = FileSizeUnit.Megabytes;
                    continue;
                }

                unsupportedSwitch = normalizedSwitch;
                return false;
            }

            return true;
        }

        private static bool TryEvaluatePage(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (candidate.LocationKind != WordFieldLocationKind.Body) {
                message = "PAGE fields outside the document body need Word layout context and were left unchanged.";
                return false;
            }

            int? page = EstimatePageForBodyField(document, candidate.AnchorElement);
            if (page == null) {
                message = "PAGE field position could not be matched to a body paragraph.";
                return false;
            }

            return TryFormatNumericField(page.Value, parsed.NumericPictureSwitch, "Updated from OfficeIMO page-break order.", out value, out status, out message);
        }

        private static bool TryFormatNumericField(
            decimal numericValue,
            string? numericPicture,
            string successMessage,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            if (!TryFormatFormulaValue(numericValue, numericPicture, out string formattedValue, out string? diagnostic)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = diagnostic == null
                    ? "Field numeric picture switch could not be applied."
                    : diagnostic.Replace("Formula numeric picture", "Field numeric picture");
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = string.IsNullOrWhiteSpace(numericPicture)
                ? successMessage
                : $"{successMessage} Numeric picture formatting was applied.";
            return true;
        }

        private enum FileSizeUnit {
            Bytes,
            Kilobytes,
            Megabytes
        }

        private static bool TrySetValue(string? source, string successMessage, out string? value, out WordFieldUpdateStatus status, out string message) {
            value = source;
            if (!string.IsNullOrEmpty(value)) {
                status = WordFieldUpdateStatus.Updated;
                message = successMessage;
                return true;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = "Source document property is empty.";
            return false;
        }

        private static bool TrySetTextValue(
            string? source,
            WordFieldInventory.ParsedFieldInstruction parsed,
            string successMessage,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            if (string.IsNullOrEmpty(source)) {
                value = null;
                status = WordFieldUpdateStatus.Skipped;
                message = "Source document property is empty.";
                return false;
            }

            string sourceText = source!;
            if (!TryApplyReferenceTextFormat(parsed.FormatSwitches, sourceText, out string formattedValue, out string? unsupportedFormat)) {
                value = null;
                status = WordFieldUpdateStatus.Unsupported;
                message = $"Field format switch {unsupportedFormat} is not supported for deterministic metadata refresh.";
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = GetLastMeaningfulFormat(parsed.FormatSwitches) == null
                ? successMessage
                : successMessage + " Text format switch was applied.";
            return true;
        }

        private static bool TrySetDate(DateTime? source, WordFieldInventory.ParsedFieldInstruction parsed, string successMessage, out string? value, out WordFieldUpdateStatus status, out string message) {
            if (!source.HasValue) {
                value = null;
                status = WordFieldUpdateStatus.Skipped;
                message = "Source document property is empty.";
                return false;
            }

            if (!TryFormatDateTime(source.Value, parsed, out value, out message)) {
                status = WordFieldUpdateStatus.Unsupported;
                return false;
            }

            if (!string.IsNullOrEmpty(value)) {
                status = WordFieldUpdateStatus.Updated;
                message = successMessage;
                return true;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = "Source document property is empty.";
            return false;
        }

        private static bool TryFormatDateTime(DateTime source, WordFieldInventory.ParsedFieldInstruction parsed, out string value, out string message) {
            string? customFormat = GetDateTimeFormatSwitch(parsed.Switches);
            if (string.IsNullOrWhiteSpace(customFormat)) {
                value = source.ToString(DefaultDateTimeFormat, CultureInfo.InvariantCulture);
                message = string.Empty;
                return true;
            }

            string normalizedFormat = NormalizeDateTimeFormat(customFormat!);
            try {
                value = source.ToString(normalizedFormat, CultureInfo.InvariantCulture);
                message = string.Empty;
                return true;
            } catch (FormatException) {
                value = string.Empty;
                message = $"Date/time format switch {customFormat} is not supported for deterministic field refresh.";
                return false;
            }
        }

        private static string? GetDateTimeFormatSwitch(IReadOnlyList<string> switches) {
            for (int index = switches.Count - 1; index >= 0; index--) {
                string fieldSwitch = switches[index].Trim();
                if (!fieldSwitch.StartsWith(@"\@", StringComparison.Ordinal)) {
                    continue;
                }

                string format = fieldSwitch.Substring(2).Trim();
                return string.IsNullOrWhiteSpace(format) ? null : TrimQuotes(format);
            }

            return null;
        }

        private static string NormalizeDateTimeFormat(string format) {
            return Regex.Replace(format, "am/pm", "tt", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static int EstimateTotalPages(WordDocument document) {
            int page = 1;
            foreach (WordParagraph paragraph in document.Paragraphs) {
                if (paragraph.IsPageBreak) {
                    page++;
                }
            }

            return page;
        }

        private static int? EstimatePageForBodyField(WordDocument document, OpenXmlElement anchorElement) {
            Body? body = document._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            Paragraph? targetParagraph = anchorElement is Paragraph paragraph
                ? paragraph
                : anchorElement.Ancestors<Paragraph>().FirstOrDefault();

            if (body == null || targetParagraph == null) {
                return null;
            }

            int page = 1;
            foreach (Paragraph currentParagraph in body.Descendants<Paragraph>()) {
                if (ReferenceEquals(currentParagraph, targetParagraph)) {
                    return page;
                }

                if (currentParagraph.Descendants<Break>().Any(documentBreak => documentBreak.Type?.Value == BreakValues.Page)) {
                    page++;
                }
            }

            return null;
        }

        private static string FormatValue(object? value) {
            if (value == null) {
                return string.Empty;
            }

            if (value is DateTime dateTime) {
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            if (value is IFormattable formattable) {
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string TrimQuotes(string value) {
            value = value.Trim();
            return value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"'
                ? value.Substring(1, value.Length - 2)
                : value;
        }

        private static void SetResultText(MutableFieldCandidate candidate, string value) {
            if (candidate.SimpleField != null) {
                List<Text> simpleTexts = candidate.SimpleField.Descendants<Text>().ToList();
                if (simpleTexts.Count > 0) {
                    SetText(simpleTexts, value);
                } else {
                    candidate.SimpleField.Append(new Run(CreateResultText(value)));
                }

                return;
            }

            var texts = candidate.ResultRuns.SelectMany(run => run.Elements<Text>()).ToList();
            if (texts.Count > 0) {
                SetText(texts, value);
                return;
            }

            var run = new Run(CreateResultText(value));
            if (candidate.EndRun != null) {
                if (!candidate.HasSeparator) {
                    candidate.EndRun.InsertBeforeSelf(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                }

                candidate.EndRun.InsertBeforeSelf(run);
            } else {
                candidate.AnchorElement.InsertAfterSelf(run);
            }

            candidate.ResultRuns.Add(run);
        }

        private static Text CreateResultText(string value) {
            var text = new Text(value);
            ApplyTextSpacePreservation(text, value);
            return text;
        }

        private static void SetText(IReadOnlyList<Text> texts, string value) {
            if (texts.Count == 0) {
                return;
            }

            texts[0].Text = value;
            ApplyTextSpacePreservation(texts[0], value);
            for (int i = 1; i < texts.Count; i++) {
                texts[i].Text = string.Empty;
                ApplyTextSpacePreservation(texts[i], string.Empty);
            }
        }

        private static void ApplyTextSpacePreservation(Text text, string value) {
            text.Space = RequiresSpacePreservation(value)
                ? SpaceProcessingModeValues.Preserve
                : null;
        }

        private static bool RequiresSpacePreservation(string value) {
            return value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]));
        }

        private static IEnumerable<MutableFieldCandidate> EnumerateFields(MainDocumentPart mainPart) {
            int sequence = 0;
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (MutableFieldCandidate candidate in EnumerateRootFields(root, ref sequence)) {
                    yield return candidate;
                }
            }
        }

        private static IEnumerable<MutableFieldCandidate> EnumerateRootFields(WordFieldInventory.FieldRoot root, ref int sequence) {
            var stack = new Stack<ComplexFieldBuilder>();
            var candidates = new List<MutableFieldCandidate>();

            foreach (OpenXmlElement element in root.Root.Descendants()) {
                if (element is SimpleField simpleField) {
                    int nestingLevel = stack.Count + simpleField.Ancestors<SimpleField>().Count();
                    candidates.Add(MutableFieldCandidate.ForSimple(root, sequence++, nestingLevel, simpleField));
                    continue;
                }

                if (element is Run run) {
                    ProcessRun(root, run, stack, candidates, ref sequence);
                }
            }

            while (stack.Count > 0) {
                ComplexFieldBuilder builder = stack.Pop();
                candidates.Add(builder.ToCandidate(root));
            }

            return candidates;
        }

        private static bool IsNestedInsideReplacedField(MutableFieldCandidate candidate, IEnumerable<MutableFieldCandidate> replacedContainingFields) {
            foreach (MutableFieldCandidate containingField in replacedContainingFields) {
                if (containingField.EndRun == null ||
                    candidate.Sequence <= containingField.Sequence ||
                    candidate.NestingLevel <= containingField.NestingLevel) {
                    continue;
                }

                if (candidate.AnchorElement.IsAfter(containingField.AnchorElement) &&
                    candidate.AnchorElement.IsBefore(containingField.EndRun)) {
                    return true;
                }
            }

            return false;
        }

        private static void ProcessRun(
            WordFieldInventory.FieldRoot root,
            Run run,
            Stack<ComplexFieldBuilder> stack,
            List<MutableFieldCandidate> candidates,
            ref int sequence) {
            FieldChar? fieldChar = run.Elements<FieldChar>().FirstOrDefault();
            FieldCharValues? fieldCharType = fieldChar?.FieldCharType?.Value;

            if (fieldCharType == FieldCharValues.Begin) {
                stack.Push(new ComplexFieldBuilder(sequence++, stack.Count, run) {
                    IsLocked = fieldChar?.FieldLock?.Value ?? false
                });
            }

            if (stack.Count == 0) {
                return;
            }

            ComplexFieldBuilder current = stack.Peek();

            foreach (FieldCode fieldCode in run.Elements<FieldCode>()) {
                current.InstructionParts.Add(InstructionPart.ForLiteral(fieldCode.Text ?? string.Empty));
            }

            if (fieldCharType == FieldCharValues.Separate) {
                current.HasSeparator = true;
            }

            if (run.Elements<Text>().Any()) {
                foreach (ComplexFieldBuilder builder in stack.Where(builder => builder.HasSeparator)) {
                    builder.ResultRuns.Add(run);
                }
            }

            if (fieldCharType == FieldCharValues.End) {
                ComplexFieldBuilder completed = stack.Pop();
                completed.EndRun = run;
                if (stack.Count > 0 && !stack.Peek().HasSeparator) {
                    ComplexFieldBuilder parent = stack.Peek();
                    completed.InstructionParentSequence = parent.Sequence;
                    parent.InstructionParts.Add(InstructionPart.ForResultRuns(completed.ResultRuns));
                }

                candidates.Add(completed.ToCandidate(root));
            }
        }

        private sealed class MutableFieldCandidate {
            private MutableFieldCandidate(
                int sequence,
                WordFieldRepresentation representation,
                WordFieldLocationKind locationKind,
                string partUri,
                IReadOnlyList<InstructionPart> instructionParts,
                int nestingLevel,
                int? instructionParentSequence,
                OpenXmlElement anchorElement,
                SimpleField? simpleField,
                List<Run> resultRuns,
                Run? endRun,
                bool hasSeparator,
                bool isLocked) {
                Sequence = sequence;
                Representation = representation;
                LocationKind = locationKind;
                PartUri = partUri;
                _instructionParts = instructionParts;
                NestingLevel = nestingLevel;
                InstructionParentSequence = instructionParentSequence;
                AnchorElement = anchorElement;
                SimpleField = simpleField;
                ResultRuns = resultRuns;
                EndRun = endRun;
                HasSeparator = hasSeparator;
                IsLocked = isLocked;
            }

            internal int Sequence { get; }

            internal int Index { get; set; }

            internal WordFieldRepresentation Representation { get; }

            internal WordFieldLocationKind LocationKind { get; }

            internal string PartUri { get; }

            private readonly IReadOnlyList<InstructionPart> _instructionParts;

            internal string InstructionText => string.Concat(_instructionParts.Select(part => part.GetText()));

            internal int NestingLevel { get; }

            internal int? InstructionParentSequence { get; }

            internal OpenXmlElement AnchorElement { get; }

            internal SimpleField? SimpleField { get; }

            internal List<Run> ResultRuns { get; }

            internal Run? EndRun { get; }

            internal bool HasSeparator { get; }

            internal bool IsLocked { get; }

            internal static MutableFieldCandidate ForSimple(WordFieldInventory.FieldRoot root, int sequence, int nestingLevel, SimpleField simpleField) {
                return new MutableFieldCandidate(
                    sequence,
                    WordFieldRepresentation.Simple,
                    root.LocationKind,
                    root.PartUri,
                    new[] { InstructionPart.ForLiteral(simpleField.Instruction?.Value ?? string.Empty) },
                    nestingLevel,
                    null,
                    simpleField,
                    simpleField,
                    simpleField.Elements<Run>().ToList(),
                    null,
                    true,
                    simpleField.FieldLock?.Value ?? false);
            }

            internal WordFieldUpdateResult ToResult(WordFieldType? fieldType, WordFieldUpdateStatus status, string? resultText, string message) {
                return new WordFieldUpdateResult(
                    Index,
                    Representation,
                    LocationKind,
                    PartUri,
                    InstructionText,
                    fieldType,
                    status,
                    resultText,
                    message);
            }

            internal static MutableFieldCandidate ForComplex(WordFieldInventory.FieldRoot root, ComplexFieldBuilder builder) {
                return new MutableFieldCandidate(
                    builder.Sequence,
                    WordFieldRepresentation.Complex,
                    root.LocationKind,
                    root.PartUri,
                    builder.InstructionParts.ToArray(),
                    builder.NestingLevel,
                    builder.InstructionParentSequence,
                    builder.AnchorElement,
                    null,
                    builder.ResultRuns,
                    builder.EndRun,
                    builder.HasSeparator,
                    builder.IsLocked);
            }
        }

        private sealed class ComplexFieldBuilder {
            internal ComplexFieldBuilder(int sequence, int nestingLevel, OpenXmlElement anchorElement) {
                Sequence = sequence;
                NestingLevel = nestingLevel;
                AnchorElement = anchorElement;
            }

            internal int Sequence { get; }

            internal int NestingLevel { get; }

            internal OpenXmlElement AnchorElement { get; }

            internal List<InstructionPart> InstructionParts { get; } = new();

            internal List<Run> ResultRuns { get; } = new();

            internal int? InstructionParentSequence { get; set; }

            internal bool HasSeparator { get; set; }

            internal bool IsLocked { get; set; }

            internal Run? EndRun { get; set; }

            internal MutableFieldCandidate ToCandidate(WordFieldInventory.FieldRoot root) {
                return MutableFieldCandidate.ForComplex(root, this);
            }
        }

        private sealed class InstructionPart {
            private readonly string? _literal;
            private readonly IReadOnlyList<Run>? _resultRuns;

            private InstructionPart(string literal) {
                _literal = literal;
            }

            private InstructionPart(IReadOnlyList<Run> resultRuns) {
                _resultRuns = resultRuns;
            }

            internal static InstructionPart ForLiteral(string literal) => new InstructionPart(literal);

            internal static InstructionPart ForResultRuns(IReadOnlyList<Run> resultRuns) => new InstructionPart(resultRuns);

            internal string GetText() {
                if (_literal != null) {
                    return _literal;
                }

                return _resultRuns == null
                    ? string.Empty
                    : string.Concat(_resultRuns.SelectMany(run => run.Elements<Text>()).Select(text => text.Text));
            }
        }

        private sealed class FieldEvaluationState {
            internal Dictionary<string, int> Sequences { get; } = new(StringComparer.OrdinalIgnoreCase);

            internal Dictionary<string, string> SequenceHeadingResetKeys { get; } = new(StringComparer.OrdinalIgnoreCase);
        }
    }
}
