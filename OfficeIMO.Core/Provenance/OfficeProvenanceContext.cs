using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal sealed class OfficeProvenanceContext {
    private readonly OfficeProvenanceOptions _options;
    private readonly Dictionary<OfficeProvenanceEvidence, string> _resolvedExternalManifestReferences =
        new Dictionary<OfficeProvenanceEvidence, string>();

    internal OfficeProvenanceContext(OfficeProvenanceAssetFormat format, OfficeProvenanceOptions options) {
        Format = format;
        _options = options;
    }

    internal OfficeProvenanceAssetFormat Format { get; }
    internal List<OfficeProvenanceEvidence> Evidence { get; } = new List<OfficeProvenanceEvidence>();
    internal List<string> Diagnostics { get; } = new List<string>();
    internal long ExpandedInspectionBytes { get; private set; }

    internal void Add(OfficeProvenanceEvidence evidence) {
        if (Evidence.Count >= _options.MaxCarriers) {
            throw OfficeProvenanceLimitException.Create($"The asset exceeds the configured carrier limit of {_options.MaxCarriers}.");
        }
        Evidence.Add(evidence);
    }

    internal void AddResolvedExternalManifestReference(OfficeProvenanceEvidence evidence, string reference) {
        if (evidence == null) throw new ArgumentNullException(nameof(evidence));
        if (string.IsNullOrWhiteSpace(reference)) throw new ArgumentException("A resolved manifest reference is required.", nameof(reference));
        _resolvedExternalManifestReferences.Add(evidence, reference);
    }

    internal void ReserveExpandedBytes(long additionalBytes, string limitMessage) {
        if (additionalBytes < 0 || ExpandedInspectionBytes > _options.MaxExpandedContainerBytes - additionalBytes) {
            throw new InvalidDataException(limitMessage);
        }
        ExpandedInspectionBytes += additionalBytes;
    }

    internal OfficeProvenanceReport ToReport() => new OfficeProvenanceReport(
        Format,
        Evidence.AsReadOnly(),
        Diagnostics.AsReadOnly(),
        ExpandedInspectionBytes,
        _resolvedExternalManifestReferences);
}
