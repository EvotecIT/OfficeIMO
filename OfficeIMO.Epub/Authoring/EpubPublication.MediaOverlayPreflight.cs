using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private EpubPreflightCheck CheckMediaOverlays(CancellationToken token) {
        var findings = new List<EpubDiagnostic>();
        EpubManifestItem[] manifest = Manifest.ToArray();
        EpubManifestItem[] overlays = manifest.Where(item => HasMediaType(item.MediaType, "application/smil+xml")).ToArray();
        if (overlays.Length == 0) return Result("media-overlays", findings);
        XElement[] metadata = OverlayDurations(Root).ToArray();
        var targetsByReference = manifest.GroupBy(item => OverlayReferenceKey(item.Reference), StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var anchors = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        bool incomplete = false;
        decimal totalTicks = 0;
        bool allDurationsKnown = true;
        foreach (EpubManifestItem overlay in overlays) {
            token.ThrowIfCancellationRequested();
            string path = overlay.Reference.ContainerPath ?? overlay.Href;
            TimeSpan? declared = null;
            try {
                XElement[] durations = metadata.Where(meta => ReferencesPackageId((string?)meta.Attribute("refines") ?? string.Empty, overlay.Id)).ToArray();
                if (durations.Length != 1) throw new InvalidDataException("Each media overlay needs exactly one media:duration refinement.");
                declared = EpubSmilClock.Parse(durations[0].Value);
                totalTicks += declared.Value.Ticks;
            } catch (Exception error) when (IsPreflightFailure(error)) {
                allDurationsKnown = false;
                Add(findings, "EPUB_PREFLIGHT_OVERLAY_DURATION", EpubDiagnosticSeverity.Error, error.Message, path);
            }
            try {
                if (overlay.Reference.Kind != EpubReferenceKind.Container || _encryption.Any(entry => entry.Path == path))
                    throw new NotSupportedException("The overlay requires external access or decryption.");
                XDocument document = GetContentXml(overlay.Id);
                if (document.Root?.Name != Smil + "smil" || (string?)document.Root.Attribute("version") != "3.0" ||
                    document.Root.Elements(Smil + "body").Count() != 1)
                    throw new InvalidDataException("A media overlay requires a SMIL 3.0 root and one body.");
                EpubContentIdentifiers.Collect(document.Root, path, true, token);
                decimal clipTicks = 0;
                bool completeClipDurations = true;
                foreach (XElement element in document.Root.Descendants()) {
                    token.ThrowIfCancellationRequested();
                    if (element.Name == Smil + "par" && (element.Elements(Smil + "text").Count() != 1 || element.Elements(Smil + "audio").Count() > 1))
                        throw new InvalidDataException("Each SMIL par requires one text target and at most one audio clip.");
                    if (element.Name == Smil + "text") {
                        if (element.Parent?.Name != Smil + "par") throw new InvalidDataException("SMIL text must belong to a par.");
                        CheckOverlayReference(overlay, path, (string?)element.Attribute("src"), false, targetsByReference, anchors, token);
                    }
                    if ((element.Name == Smil + "seq" || element.Name == Smil + "body") && element.Attribute(Ops + "textref") is XAttribute textref)
                        CheckOverlayReference(overlay, path, textref.Value, false, targetsByReference, anchors, token);
                    if (element.Name != Smil + "audio") continue;
                    if (element.Parent?.Name != Smil + "par") throw new InvalidDataException("SMIL audio must belong to a par.");
                    CheckOverlayReference(overlay, path, (string?)element.Attribute("src"), true, targetsByReference, anchors, token);
                    TimeSpan begin = element.Attribute("clipBegin") is XAttribute start ? EpubSmilClock.Parse(start.Value) : TimeSpan.Zero;
                    if (element.Attribute("clipEnd") is not XAttribute end) {
                        completeClipDurations = false;
                        continue;
                    }
                    TimeSpan finish = EpubSmilClock.Parse(end.Value);
                    if (finish <= begin) throw new InvalidDataException("An audio clip end must be later than its nonnegative start.");
                    clipTicks += finish.Ticks - begin.Ticks;
                }
                if (!completeClipDurations) {
                    incomplete = true;
                    Add(findings, "EPUB_PREFLIGHT_OVERLAY_DURATION_UNCHECKED", EpubDiagnosticSeverity.Warning,
                        "At least one audio clip omits clipEnd. Its duration requires audio decoding; clip totals were not verified.", path);
                } else if (declared.HasValue && Math.Abs(clipTicks - declared.Value.Ticks) > TimeSpan.TicksPerSecond)
                    throw new InvalidDataException("Media-overlay duration differs from its explicit audio clip sum by more than one second.");
            } catch (NotSupportedException error) {
                incomplete = true;
                Add(findings, "EPUB_PREFLIGHT_OVERLAY_UNCHECKED", EpubDiagnosticSeverity.Warning, error.Message, path);
            } catch (Exception error) when (IsPreflightFailure(error)) {
                Add(findings, "EPUB_PREFLIGHT_OVERLAY_INVALID", EpubDiagnosticSeverity.Error, error.Message, path);
            }
        }
        try {
            XElement[] totals = metadata.Where(meta => meta.Attribute("refines") == null).ToArray();
            if (totals.Length != 1) throw new InvalidDataException("A narrated publication needs exactly one total media:duration.");
            TimeSpan declaredTotal = EpubSmilClock.Parse(totals[0].Value);
            if (allDurationsKnown && Math.Abs(totalTicks - declaredTotal.Ticks) > TimeSpan.TicksPerSecond)
                throw new InvalidDataException("Total media duration differs from the overlay duration sum by more than one second.");
        } catch (Exception error) when (IsPreflightFailure(error)) {
            Add(findings, "EPUB_PREFLIGHT_OVERLAY_TOTAL_DURATION", EpubDiagnosticSeverity.Error, error.Message, PackagePath);
        }
        EpubPreflightStatus status = findings.Any(item => item.Severity == EpubDiagnosticSeverity.Error) ? EpubPreflightStatus.Failed :
            incomplete ? EpubPreflightStatus.NotChecked : EpubPreflightStatus.Passed;
        return new EpubPreflightCheck("media-overlays", status, findings);
    }

    private void CheckOverlayReference(EpubManifestItem overlay, string overlayPath, string? value, bool audio,
        IReadOnlyDictionary<string, EpubManifestItem[]> targetsByReference, Dictionary<string, HashSet<string>> anchors, CancellationToken token) {
        if (string.IsNullOrWhiteSpace(value)) throw new InvalidDataException("A media-overlay reference is missing its target.");
        EpubReference reference = EpubReference.Resolve(overlayPath, value!);
        if (!reference.IsConforming || reference.Kind != EpubReferenceKind.Container && !(audio && reference.Kind == EpubReferenceKind.External))
            throw new InvalidDataException("Invalid media-overlay reference: " + value);
        if (!targetsByReference.TryGetValue(OverlayReferenceKey(reference), out EpubManifestItem[]? targets) || targets.Any(item => audio
                ? !HasMediaType(item.MediaType, "audio/mpeg") && !HasMediaType(item.MediaType, "audio/mp4")
                : !IsSupportedContentDocument(item.MediaType) || item.MediaOverlayId != overlay.Id))
            throw new InvalidDataException("Media-overlay target lacks a matching manifest declaration and association: " + value);
        if (reference.Kind == EpubReferenceKind.External) return; // Retrieval and decoding remain explicitly unchecked.
        string path = reference.ContainerPath!;
        if (!_entries.TryGetValue(path, out byte[]? payload) || payload.Length == 0)
            throw new InvalidDataException("Media-overlay target payload is missing or empty: " + value);
        if (_encryption.Any(entry => entry.Path == path)) throw new NotSupportedException("Media-overlay target requires decryption: " + value);
        if (audio || reference.Fragment == null) return;
        if (!anchors.TryGetValue(path, out HashSet<string>? ids)) {
            token.ThrowIfCancellationRequested();
            XDocument target = ParseXml(payload, _maximumEntryBytes);
            ids = EpubContentIdentifiers.Collect(target.Root!, path, true, token);
            anchors.Add(path, ids);
        }
        if (!ids.Contains(reference.Fragment)) throw new InvalidDataException("Media-overlay text fragment is missing: " + value);
    }

    private static string OverlayReferenceKey(EpubReference reference) => reference.Kind + ":" +
        (reference.Kind == EpubReferenceKind.Container ? reference.ContainerPath : reference.ExternalUri);
}
