using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.CustomProperties;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyDocumentInfo(WordDocument source, RtfDocument destination) {
        destination.Info.Title = source.BuiltinDocumentProperties.Title;
        destination.Info.Subject = source.BuiltinDocumentProperties.Subject;
        destination.Info.Author = source.BuiltinDocumentProperties.Creator;
        destination.Info.Category = source.BuiltinDocumentProperties.Category;
        destination.Info.Keywords = source.BuiltinDocumentProperties.Keywords;
        destination.Info.Comments = source.BuiltinDocumentProperties.Description;
        destination.Info.Operator = source.BuiltinDocumentProperties.LastModifiedBy;
        destination.Info.Created = source.BuiltinDocumentProperties.Created;
        destination.Info.Revised = source.BuiltinDocumentProperties.Modified;
        destination.Info.Printed = source.BuiltinDocumentProperties.LastPrinted;
        destination.Info.Company = EmptyToNull(source.ApplicationProperties.Company);
        destination.Info.Manager = EmptyToNull(source.ApplicationProperties.Manager?.Text);
        destination.Info.HyperlinkBase = EmptyToNull(source.ApplicationProperties.HyperlinkBase?.Text);
        destination.Info.NumberOfPages = ParseInvariantInt(source.ApplicationProperties.Pages);
        destination.Info.NumberOfCharacters = ParseInvariantInt(source.ApplicationProperties.Characters);
        destination.Info.NumberOfCharactersWithSpaces = ParseInvariantInt(source.ApplicationProperties.CharactersWithSpaces);
    }

    private static void CopyCustomMetadata(WordDocument source, RtfDocument destination) {
        foreach (KeyValuePair<string, WordCustomProperty> pair in source.CustomDocumentProperties.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            destination.AddUserProperty(ToRtfUserProperty(pair.Key, pair.Value));
        }

        foreach (KeyValuePair<string, string> pair in source.DocumentVariables.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            destination.AddDocumentVariable(pair.Key, pair.Value);
        }
    }

    private static void ApplyDocumentInfo(RtfDocument source, WordDocument destination) {
        destination.BuiltinDocumentProperties.Title = source.Info.Title;
        destination.BuiltinDocumentProperties.Subject = source.Info.Subject;
        destination.BuiltinDocumentProperties.Creator = source.Info.Author;
        destination.BuiltinDocumentProperties.Category = source.Info.Category;
        destination.BuiltinDocumentProperties.Keywords = source.Info.Keywords;
        destination.BuiltinDocumentProperties.Description = source.Info.Comments;
        destination.BuiltinDocumentProperties.LastModifiedBy = source.Info.Operator;
        destination.BuiltinDocumentProperties.Created = source.Info.Created;
        destination.BuiltinDocumentProperties.Modified = source.Info.Revised;
        destination.BuiltinDocumentProperties.LastPrinted = source.Info.Printed;

        if (!string.IsNullOrEmpty(source.Info.Company)) {
            destination.ApplicationProperties.Company = source.Info.Company!;
        }

        string? manager = source.Info.Manager;
        if (!string.IsNullOrEmpty(manager)) {
            destination.ApplicationProperties.Manager = new Manager { Text = manager! };
        }

        string? hyperlinkBase = source.Info.HyperlinkBase;
        if (!string.IsNullOrEmpty(hyperlinkBase)) {
            destination.ApplicationProperties.HyperlinkBase = new HyperlinkBase { Text = hyperlinkBase! };
        }

        if (source.Info.NumberOfPages.HasValue) {
            destination.ApplicationProperties.Pages = source.Info.NumberOfPages.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (source.Info.NumberOfCharacters.HasValue) {
            destination.ApplicationProperties.Characters = source.Info.NumberOfCharacters.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (source.Info.NumberOfCharactersWithSpaces.HasValue) {
            destination.ApplicationProperties.CharactersWithSpaces = source.Info.NumberOfCharactersWithSpaces.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static void ApplyCustomMetadata(RtfDocument source, WordDocument destination) {
        foreach (RtfUserProperty property in source.UserProperties) {
            string? value = property.StaticValue ?? property.LinkedValue;
            if (value == null) continue;

            destination.CustomDocumentProperties[property.Name] = ToWordCustomProperty(property.TypeCode, value);
        }

        foreach (RtfDocumentVariable variable in source.DocumentVariables) {
            destination.SetDocumentVariable(variable.Name, variable.Value);
        }
    }

    private static RtfUserProperty ToRtfUserProperty(string name, WordCustomProperty property) {
        if (property == null) throw new ArgumentNullException(nameof(property));

        switch (property.PropertyType) {
            case PropertyTypes.DateTime:
                if (property.Value is DateTime dateTime) {
                    return RtfUserProperty.DateTime(name, dateTime);
                }

                return new RtfUserProperty(name, RtfUserProperty.DateTimeType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case PropertyTypes.NumberInteger:
                return new RtfUserProperty(name, RtfUserProperty.IntegerType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case PropertyTypes.NumberDouble:
                return new RtfUserProperty(name, RtfUserProperty.NumberType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case PropertyTypes.YesNo:
                if (property.Value is bool boolean) {
                    return RtfUserProperty.Boolean(name, boolean);
                }

                return new RtfUserProperty(name, RtfUserProperty.BooleanType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case PropertyTypes.Text:
            default:
                return new RtfUserProperty(name, RtfUserProperty.TextType, Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }
    }

    private static WordCustomProperty ToWordCustomProperty(int? typeCode, string value) {
        switch (typeCode) {
            case RtfUserProperty.IntegerType:
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int integer)) {
                    return new WordCustomProperty(integer);
                }

                break;
            case RtfUserProperty.NumberType:
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                    return new WordCustomProperty(number);
                }

                break;
            case RtfUserProperty.BooleanType:
                if (TryParseRtfBoolean(value, out bool boolean)) {
                    return new WordCustomProperty(boolean);
                }

                break;
            case RtfUserProperty.DateTimeType:
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime dateTime)) {
                    return new WordCustomProperty(dateTime);
                }

                break;
        }

        return new WordCustomProperty(value);
    }

    private static bool TryParseRtfBoolean(string value, out bool result) {
        if (string.Equals(value, "1", StringComparison.Ordinal)) {
            result = true;
            return true;
        }

        if (string.Equals(value, "0", StringComparison.Ordinal)) {
            result = false;
            return true;
        }

        return bool.TryParse(value, out result);
    }

    private static string? EmptyToNull(string? value) {
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private static int? ParseInvariantInt(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;
    }
}
