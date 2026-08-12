using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal sealed class OfficeProvenanceContext {
    private readonly OfficeProvenanceOptions _options;

    internal OfficeProvenanceContext(OfficeProvenanceAssetFormat format, OfficeProvenanceOptions options) {
        Format = format;
        _options = options;
    }

    internal OfficeProvenanceAssetFormat Format { get; }
    internal List<OfficeProvenanceEvidence> Evidence { get; } = new List<OfficeProvenanceEvidence>();
    internal List<string> Diagnostics { get; } = new List<string>();

    internal void Add(OfficeProvenanceEvidence evidence) {
        if (Evidence.Count >= _options.MaxCarriers) {
            throw new InvalidDataException($"The asset exceeds the configured carrier limit of {_options.MaxCarriers}.");
        }
        Evidence.Add(evidence);
    }

    internal OfficeProvenanceReport ToReport() => new OfficeProvenanceReport(Format, Evidence.AsReadOnly(), Diagnostics.AsReadOnly());
}
