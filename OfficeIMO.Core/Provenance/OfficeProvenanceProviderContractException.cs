using System;
using System.IO;

namespace OfficeIMO.Provenance;

/// <summary>Marks malformed optional-provider results without changing the public exception contract.</summary>
internal static class OfficeProvenanceProviderContractException {
    private const string Marker = "OfficeIMO.Provenance.ProviderContract";

    internal static InvalidDataException Create(string message) {
        var exception = new InvalidDataException(message);
        exception.Data[Marker] = true;
        return exception;
    }

    internal static bool Is(Exception exception) =>
        exception is InvalidDataException && exception.Data.Contains(Marker);
}
