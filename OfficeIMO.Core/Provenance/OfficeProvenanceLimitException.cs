using System;
using System.IO;

namespace OfficeIMO.Provenance;

/// <summary>Marks configured provenance resource limits without changing the public exception contract.</summary>
internal static class OfficeProvenanceLimitException {
    private const string Marker = "OfficeIMO.Provenance.ResourceLimit";
    private const string OutputMarker = "OfficeIMO.Provenance.OutputLimit";

    internal static InvalidDataException Create(string message) {
        var exception = new InvalidDataException(message);
        exception.Data[Marker] = true;
        return exception;
    }

    internal static InvalidDataException CreateOutput(string message) {
        InvalidDataException exception = Create(message);
        exception.Data[OutputMarker] = true;
        return exception;
    }

    internal static bool Is(Exception exception) =>
        exception is InvalidDataException && exception.Data.Contains(Marker);

    internal static bool IsOutput(Exception exception) =>
        exception is InvalidDataException && exception.Data.Contains(OutputMarker);
}
