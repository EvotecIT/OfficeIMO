using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.CustomProperties;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyDocumentInfo(WordDocument source, RtfDocument destination) {
        destination.Properties.Title = source.Properties.Title;
        destination.Properties.Subject = source.Properties.Subject;
        destination.Properties.Author = source.Properties.Creator;
        destination.Properties.Category = source.Properties.Category;
        destination.Properties.Keywords = source.Properties.Keywords;
        destination.Properties.Comments = source.Properties.Description;
        destination.Properties.Operator = source.Properties.LastModifiedBy;
        destination.Properties.Created = source.Properties.Created;
        destination.Properties.Revised = source.Properties.Modified;
        destination.Properties.Printed = source.Properties.LastPrinted;
        destination.Properties.Company = EmptyToNull(source.ApplicationProperties.Company);
        destination.Properties.Manager = EmptyToNull(source.ApplicationProperties.Manager);
        destination.Properties.HyperlinkBase = EmptyToNull(source.ApplicationProperties.HyperlinkBase);
        destination.Properties.NumberOfPages = ParseInvariantInt(source.ApplicationProperties.Pages);
        destination.Properties.NumberOfCharacters = ParseInvariantInt(source.ApplicationProperties.Characters);
        destination.Properties.NumberOfCharactersWithSpaces = ParseInvariantInt(source.ApplicationProperties.CharactersWithSpaces);
    }

    private static void CopyCustomMetadata(WordDocument source, RtfDocument destination) {
        foreach (KeyValuePair<string, WordCustomProperty> pair in source.CustomProperties.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            destination.AddUserProperty(ToRtfUserProperty(pair.Key, pair.Value));
        }

        foreach (KeyValuePair<string, string> pair in source.DocumentVariables.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            destination.AddDocumentVariable(pair.Key, pair.Value);
        }
    }

    private static void ApplyDocumentInfo(RtfDocument source, WordDocument destination) {
        destination.Properties.Title = source.Properties.Title;
        destination.Properties.Subject = source.Properties.Subject;
        destination.Properties.Creator = source.Properties.Author;
        destination.Properties.Category = source.Properties.Category;
        destination.Properties.Keywords = source.Properties.Keywords;
        destination.Properties.Description = source.Properties.Comments;
        destination.Properties.LastModifiedBy = source.Properties.Operator;
        destination.Properties.Created = source.Properties.Created;
        destination.Properties.Modified = source.Properties.Revised;
        destination.Properties.LastPrinted = source.Properties.Printed;

        if (!string.IsNullOrEmpty(source.Properties.Company)) {
            destination.ApplicationProperties.Company = source.Properties.Company!;
        }

        string? manager = source.Properties.Manager;
        if (!string.IsNullOrEmpty(manager)) {
            destination.ApplicationProperties.Manager = manager;
        }

        string? hyperlinkBase = source.Properties.HyperlinkBase;
        if (!string.IsNullOrEmpty(hyperlinkBase)) {
            destination.ApplicationProperties.HyperlinkBase = hyperlinkBase;
        }

        if (source.Properties.NumberOfPages.HasValue) {
            destination.ApplicationProperties.Pages = source.Properties.NumberOfPages.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (source.Properties.NumberOfCharacters.HasValue) {
            destination.ApplicationProperties.Characters = source.Properties.NumberOfCharacters.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (source.Properties.NumberOfCharactersWithSpaces.HasValue) {
            destination.ApplicationProperties.CharactersWithSpaces = source.Properties.NumberOfCharactersWithSpaces.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static void ApplyCustomMetadata(RtfDocument source, WordDocument destination) {
        foreach (RtfUserProperty property in source.UserProperties) {
            string? value = property.StaticValue ?? property.LinkedValue;
            if (value == null) continue;

            destination.CustomProperties[property.Name] = ToWordCustomProperty(property.TypeCode, value);
        }

        foreach (RtfDocumentVariable variable in source.DocumentVariables) {
            destination.SetDocumentVariable(variable.Name, variable.Value);
        }
    }

    private static RtfUserProperty ToRtfUserProperty(string name, WordCustomProperty property) {
        if (property == null) throw new ArgumentNullException(nameof(property));

        switch (property.PropertyType) {
            case WordCustomPropertyType.DateTime:
                if (property.Value is DateTime dateTime) {
                    return RtfUserProperty.DateTime(name, dateTime);
                }

                return new RtfUserProperty(name, RtfUserProperty.DateTimeType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case WordCustomPropertyType.NumberInteger:
                return new RtfUserProperty(name, RtfUserProperty.IntegerType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case WordCustomPropertyType.NumberDouble:
                return new RtfUserProperty(name, RtfUserProperty.NumberType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case WordCustomPropertyType.YesNo:
                if (property.Value is bool boolean) {
                    return RtfUserProperty.Boolean(name, boolean);
                }

                return new RtfUserProperty(name, RtfUserProperty.BooleanType, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
            case WordCustomPropertyType.Text:
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
