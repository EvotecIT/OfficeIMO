using System;
using System.IO;

namespace OfficeIMO.Pdf;

/// <summary>Creates and identifies invalid-data failures caused specifically by a configured PDF output limit.</summary>
internal static class PdfOutputLimitErrors {
    private static readonly object Marker = new();

    internal static InvalidDataException Create(string message) {
        InvalidDataException exception = new(message);
        exception.Data[Marker] = true;
        return exception;
    }

    internal static bool IsOutputLimitExceeded(Exception exception) =>
        exception is InvalidDataException invalidDataException && invalidDataException.Data.Contains(Marker);
}
