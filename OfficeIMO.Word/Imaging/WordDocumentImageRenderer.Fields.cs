using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static string ResolveImageExportText(
            WordParagraph paragraph,
            WordImageFlowContext? context,
            List<OfficeImageExportDiagnostic>? diagnostics = null) {
            string text;
            if (context != null && TryResolvePageFieldText(paragraph, context, out string? fieldText)) {
                text = fieldText ?? string.Empty;
            } else if (TryResolveDocumentMetadataFieldText(paragraph, out string? documentFieldText)) {
                text = documentFieldText ?? string.Empty;
            } else {
                text = paragraph.Text ?? string.Empty;
            }

            WordCapsStyle capsStyle = ResolveImageExportCapsStyle(paragraph);
            if (capsStyle == WordCapsStyle.None || text.Length == 0) {
                return text;
            }

            CultureInfo culture = CultureInfo.InvariantCulture;
            string? language = ResolveImageExportLanguage(paragraph);
            if (!string.IsNullOrWhiteSpace(language)) {
                try {
                    culture = CultureInfo.GetCultureInfo(language!);
                } catch (CultureNotFoundException) {
                    culture = CultureInfo.InvariantCulture;
                }
            }

            if (capsStyle == WordCapsStyle.SmallCaps && diagnostics != null &&
                !diagnostics.Exists(item => string.Equals(item.Code, WordImageExportDiagnosticCodes.LimitedSmallCaps, StringComparison.Ordinal))) {
                AddDiagnostic(
                    diagnostics,
                    WordImageExportDiagnosticCodes.LimitedSmallCaps,
                    "Rendered Word small-caps text as uppercase glyphs at one font size because the shared image renderer does not vary lowercase glyph sizes within a run.");
            }

            return culture.TextInfo.ToUpper(text);
        }

        private static string? ResolveImageExportLanguage(WordParagraph paragraph) {
            RunProperties? direct = paragraph.IsHyperLink ? paragraph.Hyperlink?._runProperties : paragraph._runProperties;
            string? language = direct?.Languages?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(language)) return language;
            foreach (StyleRunProperties properties in EnumerateRunStyleProperties(paragraph)) {
                language = properties.Languages?.Val?.Value;
                if (!string.IsNullOrWhiteSpace(language)) return language;
            }

            return ResolveDocumentDefaultRunProperties(paragraph)?.Languages?.Val?.Value;
        }

        private static WordCapsStyle ResolveImageExportCapsStyle(WordParagraph paragraph) {
            RunProperties? direct = paragraph.IsHyperLink ? paragraph.Hyperlink?._runProperties : paragraph._runProperties;
            bool caps = ResolveImageExportOnOff(direct?.Caps, paragraph);
            bool smallCaps = ResolveImageExportOnOff(direct?.SmallCaps, paragraph);
            return caps ? WordCapsStyle.Caps : smallCaps ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
        }

        private static bool ResolveImageExportOnOff<T>(
            T? direct,
            WordParagraph paragraph) where T : OnOffType {
            bool? directValue = ReadOnOff(direct);
            if (directValue.HasValue) return directValue.Value;
            foreach (StyleRunProperties properties in EnumerateRunStyleProperties(paragraph)) {
                bool? styleValue = ReadOnOff(properties.GetFirstChild<T>());
                if (styleValue.HasValue) return styleValue.Value;
            }

            RunPropertiesBaseStyle? defaults = ResolveDocumentDefaultRunProperties(paragraph);
            bool? defaultValue = defaults == null ? null : ReadOnOff(defaults.GetFirstChild<T>());
            if (defaultValue.HasValue) return defaultValue.Value;

            return false;
        }

        private static RunPropertiesBaseStyle? ResolveDocumentDefaultRunProperties(WordParagraph paragraph) =>
            paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?
                .DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;

        private static bool TryResolvePageFieldText(WordParagraph paragraph, WordImageFlowContext context, out string? text) {
            text = null;
            if (!context.ResolveDynamicPageFields) {
                return false;
            }

            string? simpleInstruction = paragraph._simpleField?.Instruction?.Value;
            if (!string.IsNullOrWhiteSpace(simpleInstruction)) {
                return TryResolvePageFieldInstruction(simpleInstruction!, context, out text);
            }

            if (TryResolveComplexFieldText(paragraph, (string instruction, out string? resolved) => TryResolvePageFieldInstruction(instruction, context, out resolved), out text)) {
                return true;
            }

            return false;
        }

        private static bool TryResolvePageFieldInstruction(string instruction, WordImageFlowContext context, out string? text) {
            text = null;
            if (IsFieldInstruction(instruction, "PAGE")) {
                text = TryReadFieldNumberFormat(instruction, out NumberFormatValues? format)
                    ? FormatPageNumber(context.PageNumberValue, format)
                    : context.PageNumberText;
                return true;
            }

            if (IsFieldInstruction(instruction, "NUMPAGES")) {
                text = TryReadFieldNumberFormat(instruction, out NumberFormatValues? format)
                    ? FormatPageNumber(context.TotalPageCount, format)
                    : context.TotalPageCount.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            if (IsFieldInstruction(instruction, "SECTION")) {
                text = TryReadFieldNumberFormat(instruction, out NumberFormatValues? format)
                    ? FormatPageNumber(context.SectionNumber, format)
                    : context.SectionNumber.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            if (IsFieldInstruction(instruction, "SECTIONPAGES")) {
                text = TryReadFieldNumberFormat(instruction, out NumberFormatValues? format)
                    ? FormatPageNumber(context.SectionPageCount, format)
                    : context.SectionPageCount.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            return false;
        }

        private delegate bool FieldInstructionResolver(string instruction, out string? text);

        private static bool TryResolveComplexFieldText(WordParagraph paragraph, FieldInstructionResolver resolver, out string? text) {
            text = null;
            if (paragraph._runs == null || paragraph._runs.Count == 0) {
                return false;
            }

            var output = new StringBuilder();
            var instruction = new StringBuilder();
            int fieldDepth = 0;
            bool inFieldResult = false;
            bool skipCurrentFieldResult = false;
            bool resolvedAnyField = false;

            foreach (Run run in paragraph._runs) {
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (child is FieldChar fieldChar) {
                        FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                        if (fieldCharType == FieldCharValues.Begin) {
                            if (fieldDepth == 0) {
                                instruction.Clear();
                                inFieldResult = false;
                                skipCurrentFieldResult = false;
                            }

                            fieldDepth++;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.Separate && fieldDepth > 0) {
                            if (fieldDepth == 1 && resolver(instruction.ToString(), out string? resolvedText)) {
                                output.Append(resolvedText);
                                skipCurrentFieldResult = true;
                                resolvedAnyField = true;
                            } else {
                                skipCurrentFieldResult = false;
                            }

                            inFieldResult = true;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.End && fieldDepth > 0) {
                            if (fieldDepth == 1 && !inFieldResult && resolver(instruction.ToString(), out string? resolvedText)) {
                                output.Append(resolvedText);
                                resolvedAnyField = true;
                            }

                            fieldDepth--;
                            if (fieldDepth == 0) {
                                instruction.Clear();
                                inFieldResult = false;
                                skipCurrentFieldResult = false;
                            }

                            continue;
                        }
                    }

                    if (fieldDepth > 0) {
                        if (!inFieldResult) {
                            if (child is FieldCode fieldCode) {
                                instruction.Append(fieldCode.Text);
                            }

                            continue;
                        }

                        if (skipCurrentFieldResult) {
                            continue;
                        }
                    }

                    AppendRenderableRunText(child, output);
                }
            }

            if (!resolvedAnyField) {
                return false;
            }

            text = output.ToString();
            return true;
        }

        private static void AppendRenderableRunText(OpenXmlElement element, StringBuilder builder) {
            switch (element) {
                case Text textElement:
                    builder.Append(textElement.Text);
                    break;
                case TabChar:
                    builder.Append('\t');
                    break;
                case Break:
                    builder.AppendLine();
                    break;
            }
        }

        private static bool TryResolveDocumentMetadataFieldText(WordParagraph paragraph, out string? text) {
            text = null;
            WordDocument? document = paragraph._document;
            if (document == null) {
                return false;
            }

            if (TryResolveComplexFieldText(paragraph, (string instruction, out string? resolved) => TryResolveDocumentMetadataFieldInstruction(document, instruction, out resolved), out text)) {
                return true;
            }

            if (!TryReadFieldInstruction(paragraph, out string? instruction)) {
                return false;
            }

            return TryResolveDocumentMetadataFieldInstruction(document, instruction!, out text);
        }

        private static bool TryResolveDocumentMetadataFieldInstruction(WordDocument document, string instruction, out string? text) {
            text = null;
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction!);
            if (parsed.Diagnostics.Count > 0 || parsed.FieldType == null) {
                return false;
            }

            switch (parsed.FieldType.Value) {
                case WordFieldType.Author:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Creator, parsed, out text);
                case WordFieldType.Title:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Title, parsed, out text);
                case WordFieldType.Subject:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Subject, parsed, out text);
                case WordFieldType.Keywords:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Keywords, parsed, out text);
                case WordFieldType.Comments:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Description, parsed, out text);
                case WordFieldType.LastSavedBy:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.LastModifiedBy, parsed, out text);
                case WordFieldType.RevNum:
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Revision, parsed, out text);
                case WordFieldType.CreateDate:
                    return TryResolveDateFieldValue(document.BuiltinDocumentProperties.Created, parsed, out text);
                case WordFieldType.PrintDate:
                    return TryResolveDateFieldValue(document.BuiltinDocumentProperties.LastPrinted, parsed, out text);
                case WordFieldType.SaveDate:
                    return TryResolveDateFieldValue(document.BuiltinDocumentProperties.Modified, parsed, out text);
                case WordFieldType.FileName:
                    return TryResolveFileNameFieldValue(document, parsed, out text);
                case WordFieldType.Info:
                case WordFieldType.DocProperty:
                    return TryResolveNamedDocumentPropertyFieldValue(document, parsed, out text);
                default:
                    return false;
            }
        }

        private static bool TryResolveNamedDocumentPropertyFieldValue(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? text) {
            text = null;
            string? propertyName = parsed.Instructions.Count > 0 ? parsed.Instructions[0] : null;
            if (string.IsNullOrWhiteSpace(propertyName)) {
                return false;
            }

            propertyName = TrimFieldQuotes(propertyName!);
            if (TryResolveBuiltInDocumentProperty(document, propertyName, parsed, out text)) {
                return true;
            }

            foreach (var pair in document.CustomDocumentProperties) {
                if (!string.Equals(pair.Key, propertyName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (pair.Value.Date.HasValue) {
                    return TryResolveDateFieldValue(pair.Value.Date.Value, parsed, out text);
                }

                return TryResolveTextFieldValue(FormatMetadataValue(pair.Value.Value), parsed, out text);
            }

            return false;
        }

        private static bool TryResolveBuiltInDocumentProperty(
            WordDocument document,
            string propertyName,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? text) {
            text = null;
            switch (propertyName.Replace(" ", string.Empty).ToUpperInvariant()) {
                case "AUTHOR":
                case "CREATOR":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Creator, parsed, out text);
                case "TITLE":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Title, parsed, out text);
                case "SUBJECT":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Subject, parsed, out text);
                case "CATEGORY":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Category, parsed, out text);
                case "KEYWORDS":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Keywords, parsed, out text);
                case "COMMENTS":
                case "DESCRIPTION":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Description, parsed, out text);
                case "LASTSAVEDBY":
                case "LASTMODIFIEDBY":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.LastModifiedBy, parsed, out text);
                case "CREATED":
                case "CREATEDATE":
                    return TryResolveDateFieldValue(document.BuiltinDocumentProperties.Created, parsed, out text);
                case "LASTPRINTED":
                case "PRINTDATE":
                    return TryResolveDateFieldValue(document.BuiltinDocumentProperties.LastPrinted, parsed, out text);
                case "MODIFIED":
                case "SAVEDATE":
                    return TryResolveDateFieldValue(document.BuiltinDocumentProperties.Modified, parsed, out text);
                case "REVISION":
                case "REVNUM":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Revision, parsed, out text);
                case "VERSION":
                    return TryResolveTextFieldValue(document.BuiltinDocumentProperties.Version, parsed, out text);
                default:
                    return false;
            }
        }

        private static bool TryResolveFileNameFieldValue(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? text) {
            text = null;
            string? filePath = document.FilePath;
            if (string.IsNullOrWhiteSpace(filePath)) {
                return false;
            }

            bool includePath = false;
            foreach (string fieldSwitch in parsed.Switches) {
                if (string.Equals(fieldSwitch.Trim(), "\\p", StringComparison.OrdinalIgnoreCase)) {
                    includePath = true;
                    break;
                }
            }

            string source = includePath ? filePath! : Path.GetFileName(filePath!);
            return TryResolveTextFieldValue(source, parsed, out text);
        }

        private static bool TryResolveTextFieldValue(
            string? source,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? text) {
            text = null;
            if (string.IsNullOrEmpty(source)) {
                return false;
            }

            if (!TryApplyMetadataTextFormat(parsed.FormatSwitches, source!, out string formatted)) {
                return false;
            }

            text = formatted;
            return true;
        }

        private static bool TryResolveDateFieldValue(
            DateTime? source,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? text) {
            text = null;
            if (!source.HasValue) {
                return false;
            }

            string value = source.Value.ToString(GetMetadataDateFormat(parsed), CultureInfo.InvariantCulture);
            if (!TryApplyMetadataTextFormat(parsed.FormatSwitches, value, out string formatted)) {
                return false;
            }

            text = formatted;
            return true;
        }

        private static string GetMetadataDateFormat(WordFieldInventory.ParsedFieldInstruction parsed) {
            for (int index = parsed.Switches.Count - 1; index >= 0; index--) {
                string fieldSwitch = parsed.Switches[index].Trim();
                if (!fieldSwitch.StartsWith("\\@", StringComparison.Ordinal)) {
                    continue;
                }

                string format = TrimFieldQuotes(fieldSwitch.Substring(2).Trim());
                if (!string.IsNullOrWhiteSpace(format)) {
                    return ReplaceMetadataDateAmPmToken(format);
                }
            }

            return "yyyy-MM-dd HH:mm:ss";
        }

        private static string ReplaceMetadataDateAmPmToken(string format) {
            const string token = "am/pm";
            var builder = new StringBuilder();
            int start = 0;
            while (start < format.Length) {
                int index = format.IndexOf(token, start, StringComparison.OrdinalIgnoreCase);
                if (index < 0) {
                    builder.Append(format, start, format.Length - start);
                    break;
                }

                builder.Append(format, start, index - start);
                builder.Append("tt");
                start = index + token.Length;
            }

            return builder.ToString();
        }

        private static bool TryApplyMetadataTextFormat(
            System.Collections.Generic.IReadOnlyList<WordFieldFormat> formatSwitches,
            string source,
            out string value) {
            value = source;
            WordFieldFormat? format = GetLastMetadataTextFormat(formatSwitches);
            switch (format) {
                case null:
                    return true;
                case WordFieldFormat.Lower:
                    value = source.ToLowerInvariant();
                    return true;
                case WordFieldFormat.Upper:
                    value = source.ToUpperInvariant();
                    return true;
                case WordFieldFormat.FirstCap:
                    value = CapitalizeMetadataFirstWord(source);
                    return true;
                case WordFieldFormat.Caps:
                    value = CapitalizeMetadataEachWord(source);
                    return true;
                default:
                    return false;
            }
        }

        private static WordFieldFormat? GetLastMetadataTextFormat(System.Collections.Generic.IReadOnlyList<WordFieldFormat> formatSwitches) {
            WordFieldFormat? format = null;
            foreach (WordFieldFormat fieldFormat in formatSwitches) {
                if (fieldFormat == WordFieldFormat.Mergeformat || fieldFormat == WordFieldFormat.CharFormat) {
                    continue;
                }

                format = fieldFormat;
            }

            return format;
        }

        private static string CapitalizeMetadataFirstWord(string source) {
            if (source.Length == 0) {
                return source;
            }

            char[] chars = source.ToLowerInvariant().ToCharArray();
            for (int index = 0; index < chars.Length; index++) {
                if (char.IsLetter(chars[index])) {
                    chars[index] = char.ToUpperInvariant(chars[index]);
                    break;
                }
            }

            return new string(chars);
        }

        private static string CapitalizeMetadataEachWord(string source) {
            if (source.Length == 0) {
                return source;
            }

            char[] chars = source.ToLowerInvariant().ToCharArray();
            bool nextLetterStartsWord = true;
            for (int index = 0; index < chars.Length; index++) {
                if (!char.IsLetter(chars[index])) {
                    nextLetterStartsWord = true;
                    continue;
                }

                if (nextLetterStartsWord) {
                    chars[index] = char.ToUpperInvariant(chars[index]);
                    nextLetterStartsWord = false;
                }
            }

            return new string(chars);
        }

        private static string FormatMetadataValue(object? value) {
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

        private static string TrimFieldQuotes(string value) {
            value = value.Trim();
            return value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"'
                ? value.Substring(1, value.Length - 2)
                : value;
        }

        private static bool TryReadFieldInstruction(WordParagraph paragraph, out string? instruction) {
            instruction = paragraph._simpleField?.Instruction?.Value;
            if (!string.IsNullOrWhiteSpace(instruction)) {
                return true;
            }

            if (paragraph._runs == null || paragraph._runs.Count == 0) {
                return false;
            }

            var builder = new StringBuilder();
            foreach (Run run in paragraph._runs) {
                foreach (FieldCode fieldCode in run.Descendants<FieldCode>()) {
                    builder.Append(fieldCode.Text);
                }
            }

            instruction = builder.ToString();
            return !string.IsNullOrWhiteSpace(instruction);
        }

        private static bool IsFieldInstruction(string instruction, string expectedKeyword) {
            string trimmed = instruction.TrimStart();
            if (trimmed.Length == 0) {
                return false;
            }

            int length = 0;
            while (length < trimmed.Length && !char.IsWhiteSpace(trimmed[length]) && trimmed[length] != '\\') {
                length++;
            }

            if (length == 0) {
                return false;
            }

            return string.Equals(trimmed.Substring(0, length), expectedKeyword, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryReadFieldNumberFormat(string instruction, out NumberFormatValues? format) {
            format = null;
            string[] tokens = instruction.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < tokens.Length - 1; i++) {
                if (!string.Equals(tokens[i], "\\*", StringComparison.Ordinal)) {
                    continue;
                }

                string switchValue = tokens[i + 1].Trim('"');
                if (string.Equals(switchValue, "MERGEFORMAT", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (string.Equals(switchValue, "ROMAN", StringComparison.Ordinal)) {
                    format = NumberFormatValues.UpperRoman;
                    return true;
                }

                if (string.Equals(switchValue, "roman", StringComparison.Ordinal)) {
                    format = NumberFormatValues.LowerRoman;
                    return true;
                }

                if (string.Equals(switchValue, "ALPHABETIC", StringComparison.Ordinal)) {
                    format = NumberFormatValues.UpperLetter;
                    return true;
                }

                if (string.Equals(switchValue, "alphabetic", StringComparison.Ordinal)) {
                    format = NumberFormatValues.LowerLetter;
                    return true;
                }

                if (string.Equals(switchValue, "ArabicZero", StringComparison.OrdinalIgnoreCase)) {
                    format = NumberFormatValues.DecimalZero;
                    return true;
                }

                if (string.Equals(switchValue, "Arabic", StringComparison.OrdinalIgnoreCase)) {
                    format = null;
                    return true;
                }
            }

            return false;
        }

        private static string FormatPageNumber(int number, NumberFormatValues? format) {
            if (format == NumberFormatValues.UpperRoman) {
                return ToRoman(number).ToUpperInvariant();
            }

            if (format == NumberFormatValues.LowerRoman) {
                return ToRoman(number).ToLowerInvariant();
            }

            if (format == NumberFormatValues.UpperLetter) {
                return ToAlphabetic(number, uppercase: true);
            }

            if (format == NumberFormatValues.LowerLetter) {
                return ToAlphabetic(number, uppercase: false);
            }

            if (format == NumberFormatValues.DecimalZero && number >= 0 && number < 10) {
                return "0" + number.ToString(CultureInfo.InvariantCulture);
            }

            return number.ToString(CultureInfo.InvariantCulture);
        }

        private static string ToAlphabetic(int number, bool uppercase) {
            if (number <= 0) {
                return number.ToString(CultureInfo.InvariantCulture);
            }

            var builder = new StringBuilder();
            int value = number;
            while (value > 0) {
                value--;
                char letter = (char)((uppercase ? 'A' : 'a') + (value % 26));
                builder.Insert(0, letter);
                value /= 26;
            }

            return builder.ToString();
        }

        private static string ToRoman(int number) {
            if (number <= 0) {
                return number.ToString(CultureInfo.InvariantCulture);
            }

            var numerals = new[] {
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
            };
            var builder = new StringBuilder();
            int remaining = number;
            foreach (var (value, symbol) in numerals) {
                while (remaining >= value) {
                    builder.Append(symbol);
                    remaining -= value;
                }
            }

            return builder.ToString();
        }
    }
}
