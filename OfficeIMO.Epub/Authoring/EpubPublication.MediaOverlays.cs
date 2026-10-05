using System.Globalization;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private const string OverlayVocabulary = "http://www.idpf.org/epub/vocab/overlays/#";
    private static readonly XNamespace Smil = "http://www.w3.org/ns/SMIL";

    /// <summary>
    /// Atomically adds sequential SMIL narration for one XHTML document, its manifest association,
    /// and exact clip-sum duration metadata. Existing overlays are never replaced. Audio durations
    /// are caller declarations; encoded audio and reading-system playback require independent validation.
    /// </summary>
    public EpubManifestItem AddMediaOverlay(string contentManifestId, string overlayId, string containerPath,
        EpubMediaOverlay overlay, CancellationToken cancellationToken = default) {
        if (overlay == null) throw new ArgumentNullException(nameof(overlay));
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Media-overlay authoring requires EPUB 3.");
        if (EpubVocabulary.Expand(Root, "media:duration") != OverlayVocabulary + "duration")
            throw new InvalidDataException("The media prefix is mapped to a different vocabulary.");
        VerifyContentPath(containerPath);
        VerifyAvailableId(overlayId);
        EpubManifestItem contentItem = RequireManifestItem(contentManifestId);
        if (!string.IsNullOrEmpty(contentItem.MediaOverlayId)) throw new InvalidOperationException("The content already has a media overlay.");
        var prepared = PrepareMediaOverlay(contentItem, containerPath, overlay, cancellationToken);
        byte[] payload = SerializeXml(prepared.Document, _maximumEntryBytes);
        XElement declaration = PrepareResourceDeclaration(overlayId, containerPath, "application/smil+xml", payload, null);
        TimeSpan duration = prepared.Duration;
        if (OverlayDurations(Root).Any(meta => ReferencesPackageId((string?)meta.Attribute("refines") ?? string.Empty, overlayId)))
            throw new InvalidDataException("The new overlay id is already the target of retained duration metadata.");
        TimeSpan total = MediaOverlayTotal(duration);
        cancellationToken.ThrowIfCancellationRequested();
        EditPackageElement(Root, proposed => {
            cancellationToken.ThrowIfCancellationRequested();
            XElement manifest = proposed.Element(Opf + "manifest")!;
            manifest.Add(new XElement(declaration));
            manifest.Elements(Opf + "item").Single(item => (string?)item.Attribute("id") == contentManifestId).SetAttributeValue("media-overlay", overlayId);
            XElement metadata = proposed.Element(Opf + "metadata")!;
            XElement? existingTotal = OverlayDurations(proposed).SingleOrDefault(meta => meta.Attribute("refines") == null);
            if (existingTotal == null) metadata.Add(DurationMetadata(total, null));
            else existingTotal.Value = EpubSmilClock.Format(total);
            metadata.Add(DurationMetadata(duration, "#" + overlayId));
        }, payload.LongLength);
        _entries.Add(containerPath, payload);
        _retainedBytes += payload.LongLength;
        return RequireManifestItem(overlayId);
    }

    private (XDocument Document, TimeSpan Duration) PrepareMediaOverlay(EpubManifestItem contentItem, string containerPath,
        EpubMediaOverlay overlay, CancellationToken cancellationToken) {
        string contentPath = RequireLocalPath(contentItem);
        if (Manifest.Count(item => item.Reference.ContainerPath == contentPath) != 1)
            throw new InvalidDataException("Media-overlay authoring requires a uniquely declared content resource.");
        XDocument content = EditableXhtml(contentItem.Id);
        XElement body = content.Root!.Element(Html + "body") ?? throw new InvalidDataException("Content has no XHTML body.");
        EpubContentIdentifiers.Collect(content.Root, contentPath, rejectDuplicates: true, cancellationToken);
        if (overlay.Cues == null || overlay.Cues.Count < 1 || overlay.Cues.Count > 10000)
            throw new ArgumentException("An overlay requires one to 10,000 cues.", nameof(overlay));
        if (overlay.AudioDurations == null || overlay.AudioDurations.Count > 10000)
            throw new ArgumentException("Audio duration declarations are required and limited to 10,000 resources.", nameof(overlay));
        var targets = body.DescendantsAndSelf().Select((element, index) => new { Element = element, Index = index })
            .Where(item => item.Element.Attribute("id") != null)
            .ToDictionary(item => (string)item.Element.Attribute("id")!, item => (item.Element, item.Index), StringComparer.Ordinal);
        var audioPaths = new Dictionary<string, string>(StringComparer.Ordinal);
        var sequence = new XElement(Smil + "seq", new XAttribute(Ops + "textref", RelativeHref(containerPath, contentPath)));
        int previousIndex = -1;
        int cueIndex = 0;
        XElement? previousTarget = null;
        long ticks = 0;
        foreach (EpubMediaOverlayCue cue in overlay.Cues) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cue == null || cue.ElementId.Length == 0 || cue.ElementId.Length > 1024 ||
                !targets.TryGetValue(cue.ElementId, out var target) || target.Element.Name.Namespace != Html ||
                target.Index <= previousIndex || previousTarget != null && target.Element.Ancestors().Contains(previousTarget))
                throw new InvalidDataException("Cue targets must be distinct, non-nested XHTML body elements in document order.");
            if (new[] { "script", "style", "link", "meta", "template" }.Contains(target.Element.Name.LocalName) ||
                target.Element.Ancestors(Html + "template").Any())
                throw new InvalidDataException("A narration cue must select body content.");
            if (!overlay.AudioDurations.TryGetValue(cue.AudioManifestId, out TimeSpan audioDuration) || audioDuration <= TimeSpan.Zero ||
                cue.ClipBegin < TimeSpan.Zero || cue.ClipEnd <= cue.ClipBegin || cue.ClipEnd > audioDuration)
                throw new ArgumentException("Each clip requires 0 <= begin < end <= the declared positive audio duration.", nameof(overlay));
            if (!audioPaths.TryGetValue(cue.AudioManifestId, out string? audioPath)) {
                EpubManifestItem audio = RequireManifestItem(cue.AudioManifestId);
                if (!HasMediaType(audio.MediaType, "audio/mpeg") && !HasMediaType(audio.MediaType, "audio/mp4"))
                    throw new NotSupportedException("Narration requires audio/mpeg or audio/mp4.");
                audioPath = RequireLocalPath(audio);
                if (!_entries.TryGetValue(audioPath, out byte[]? data) || data.Length == 0 || _encryption.Any(item => item.Path == audioPath))
                    throw new NotSupportedException("Narration requires retained, nonempty, unencrypted audio.");
                audioPaths.Add(cue.AudioManifestId, audioPath);
            }
            try { ticks = checked(ticks + (cue.ClipEnd.Ticks - cue.ClipBegin.Ticks)); }
            catch (OverflowException) { throw new ArgumentException("Overlay duration exceeds TimeSpan.", nameof(overlay)); }
            sequence.Add(new XElement(Smil + "par", new XAttribute("id", "cue" + (cueIndex++).ToString(CultureInfo.InvariantCulture)),
                new XElement(Smil + "text", new XAttribute("src", RelativeHref(containerPath, contentPath) + "#" + Uri.EscapeDataString(cue.ElementId))),
                new XElement(Smil + "audio", new XAttribute("src", RelativeHref(containerPath, audioPath)),
                    new XAttribute("clipBegin", EpubSmilClock.Format(cue.ClipBegin)), new XAttribute("clipEnd", EpubSmilClock.Format(cue.ClipEnd)))));
            previousIndex = target.Index; previousTarget = target.Element;
        }
        var document = new XDocument(new XElement(Smil + "smil", new XAttribute("version", "3.0"),
            new XAttribute(XNamespace.Xmlns + "epub", Ops.NamespaceName), new XElement(Smil + "body", sequence)));
        return (document, TimeSpan.FromTicks(ticks));
    }

    private TimeSpan MediaOverlayTotal(TimeSpan addedDuration, string? replacedOverlayId = null) {
        XElement[] durations = OverlayDurations(Root).ToArray();
        if (durations.Count(meta => meta.Attribute("refines") == null) > 1)
            throw new InvalidDataException("The publication has ambiguous total media durations.");
        long total = addedDuration.Ticks;
        foreach (EpubManifestItem item in Manifest.Where(item => HasMediaType(item.MediaType, "application/smil+xml") && item.Id != replacedOverlayId)) {
            XElement[] values = durations.Where(meta => ReferencesPackageId((string?)meta.Attribute("refines") ?? string.Empty, item.Id)).ToArray();
            if (values.Length != 1) throw new InvalidDataException("Each existing SMIL resource requires exactly one duration before changing narration.");
            try { total = checked(total + EpubSmilClock.Parse(values[0].Value).Ticks); }
            catch (OverflowException) { throw new InvalidDataException("Total narration duration exceeds TimeSpan."); }
        }
        return TimeSpan.FromTicks(total);
    }

    private static IEnumerable<XElement> OverlayDurations(XElement root) => root.Element(Opf + "metadata")!.Elements(Opf + "meta")
        .Where(meta => EpubVocabulary.Expand(root, (string?)meta.Attribute("property") ?? string.Empty) == OverlayVocabulary + "duration");
    private static XElement DurationMetadata(TimeSpan duration, string? refines) => new XElement(Opf + "meta",
        new XAttribute("property", "media:duration"), refines == null ? null : new XAttribute("refines", refines), EpubSmilClock.Format(duration));
}
