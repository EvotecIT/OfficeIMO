using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Visio {
    internal static class VisioStencilMigrationPlanTextSerializer {
        internal static VisioStencilMigrationPlan FromText(string text) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            Dictionary<string, string> values = ParseLines(text);
            int count = ReadRequiredInt(values, "migration.count");
            List<VisioStencilMigrationPlannedReplacement> replacements = new(count);

            for (int i = 0; i < count; i++) {
                string prefix = "migration.replacement[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                replacements.Add(new VisioStencilMigrationPlannedReplacement(
                    ReadNullableInt(values, prefix + ".pageId"),
                    ReadNullableString(values, prefix + ".page", prefix + ".page.isNull"),
                    ReadNullableString(values, prefix + ".pageNameU", prefix + ".pageNameU.isNull"),
                    ReadRequiredString(values, prefix + ".shapeId"),
                    ReadNullableString(values, prefix + ".text", prefix + ".text.isNull"),
                    ReadRequiredEnum(values, prefix + ".matchKind"),
                    ReadNullableString(values, prefix + ".matchValue", prefix + ".matchValue.isNull"),
                    ReadNullableString(values, prefix + ".oldMasterNameU", prefix + ".oldMasterNameU.isNull"),
                    ReadNullableString(values, prefix + ".newMasterNameU", prefix + ".newMasterNameU.isNull"),
                    ReadNullableString(values, prefix + ".oldStencilId", prefix + ".oldStencilId.isNull"),
                    ReadNullableString(values, prefix + ".newStencilId", prefix + ".newStencilId.isNull"),
                    ReadRequiredString(values, prefix + ".replacementStencilName"),
                    ReadRequiredString(values, prefix + ".replacementStencilCategory"),
                    ReadRequiredBool(values, prefix + ".resizeToStencil")));
            }

            return new VisioStencilMigrationPlan(replacements);
        }

        private static Dictionary<string, string> ParseLines(string text) {
            Dictionary<string, string> values = new(StringComparer.Ordinal);
            using StringReader reader = new(text);
            string? line;
            int lineNumber = 0;
            while ((line = reader.ReadLine()) != null) {
                lineNumber++;
                if (line.Length == 0) {
                    continue;
                }

                int separator = line.IndexOf('=');
                if (separator <= 0) {
                    throw new InvalidDataException($"Migration plan artifact line {lineNumber.ToString(CultureInfo.InvariantCulture)} is not a key-value pair.");
                }

                string key = line.Substring(0, separator);
                string value = VisioInspectionSnapshot.UnescapeValue(line.Substring(separator + 1));
                values[key] = value;
            }

            return values;
        }

        private static string ReadRequiredString(Dictionary<string, string> values, string key) {
            if (!values.TryGetValue(key, out string? value)) {
                throw new InvalidDataException($"Migration plan artifact is missing required key '{key}'.");
            }

            return value;
        }

        private static string? ReadNullableString(Dictionary<string, string> values, string key, string nullKey) {
            if (values.TryGetValue(nullKey, out string? isNullText) && ParseBool(isNullText, nullKey)) {
                return null;
            }

            if (!values.TryGetValue(key, out string? value)) {
                return null;
            }

            return value;
        }

        private static int ReadRequiredInt(Dictionary<string, string> values, string key) {
            string value = ReadRequiredString(values, key);
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result) || result < 0) {
                throw new InvalidDataException($"Migration plan artifact key '{key}' is not a valid non-negative integer.");
            }

            return result;
        }

        private static int? ReadNullableInt(Dictionary<string, string> values, string key) {
            if (values.TryGetValue(key + ".hasValue", out string? hasValueText) && !ParseBool(hasValueText, key + ".hasValue")) {
                return null;
            }

            if (!values.TryGetValue(key, out string? value) || string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)) {
                throw new InvalidDataException($"Migration plan artifact key '{key}' is not a valid integer.");
            }

            return result;
        }

        private static bool ReadRequiredBool(Dictionary<string, string> values, string key) {
            return ParseBool(ReadRequiredString(values, key), key);
        }

        private static bool ParseBool(string value, string key) {
            if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            throw new InvalidDataException($"Migration plan artifact key '{key}' is not a valid Boolean value.");
        }

        private static VisioStencilMigrationMatchKind ReadRequiredEnum(Dictionary<string, string> values, string key) {
            string value = ReadRequiredString(values, key);
            if (Enum.TryParse(value, ignoreCase: false, out VisioStencilMigrationMatchKind result) &&
                Enum.IsDefined(typeof(VisioStencilMigrationMatchKind), result)) {
                return result;
            }

            throw new InvalidDataException($"Migration plan artifact key '{key}' is not a valid migration match kind.");
        }
    }
}
