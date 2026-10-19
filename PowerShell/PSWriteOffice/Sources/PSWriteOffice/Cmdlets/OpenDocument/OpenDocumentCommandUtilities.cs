using System;
using System.Management.Automation;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

internal static class OpenDocumentCommandUtilities
{
    internal static void ValidateOpenDocumentExtension(string path, OdfDocumentKind kind, string parameterName)
    {
        var expected = kind switch
        {
            OdfDocumentKind.Text => ".odt",
            OdfDocumentKind.Spreadsheet => ".ods",
            OdfDocumentKind.Presentation => ".odp",
            OdfDocumentKind.Graphics => ".odg",
            _ => throw new InvalidOperationException("Unsupported OpenDocument kind.")
        };
        ValidateExtension(path, expected, kind, parameterName);
    }

    internal static void ValidateOfficeExtension(string path, OdfDocumentKind kind, string parameterName)
    {
        var expected = kind switch
        {
            OdfDocumentKind.Text => ".docx",
            OdfDocumentKind.Spreadsheet => ".xlsx",
            OdfDocumentKind.Presentation => ".pptx",
            _ => throw new InvalidOperationException("Unsupported OpenDocument kind.")
        };
        ValidateExtension(path, expected, kind, parameterName);
    }

    private static void ValidateExtension(string path, string expected, OdfDocumentKind kind, string parameterName)
    {
        var actual = System.IO.Path.GetExtension(path);
        if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException($"{parameterName} must use the {expected} extension for {kind} content.", parameterName);
        }
    }
}
