using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Tests;

internal static class OpenXmlValidationFormatting {
    public static string FormatValidationErrors(IEnumerable<ValidationErrorInfo> errors) {
        return string.Join(Environment.NewLine + Environment.NewLine,
            errors.Select(error =>
                $"Description: {error.Description}\n" +
                $"Id: {error.Id}\n" +
                $"ErrorType: {error.ErrorType}\n" +
                $"Part: {error.Part?.Uri}\n" +
                $"Path: {error.Path?.XPath}"));
    }
}
