using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Workflows;

public static partial class OfficeRasterContentSafety {
    private static AnalysisState Analyze(
        OfficeRasterImage image,
        OcrResult result,
        string engineId,
        OfficeRasterContentSafetyOptions.Snapshot options,
        RasterWorkBudget budget,
        CancellationToken cancellationToken) {
        if (result == null) throw new InvalidDataException("The OCR engine returned no result.");
        IReadOnlyList<OcrTextSpan> rawSpans = result.Spans ?? Array.Empty<OcrTextSpan>();
        IReadOnlyList<OcrDiagnostic> rawDiagnostics = result.Diagnostics ?? Array.Empty<OcrDiagnostic>();
        if (rawSpans.Count > options.MaximumOcrSpans) {
            throw new InvalidDataException("The OCR result exceeds the configured span limit.");
        }
        if (rawDiagnostics.Count > options.MaximumOcrSpans) {
            throw new InvalidDataException("The OCR result exceeds the configured diagnostic limit.");
        }
        ValidateOcrOutputCharacters(result, rawSpans, rawDiagnostics, options, cancellationToken);
        if (rawDiagnostics.Any(diagnostic => diagnostic != null &&
            !Enum.IsDefined(typeof(OcrDiagnosticSeverity), diagnostic.Severity))) {
            throw new InvalidDataException("OCR reported an undefined diagnostic severity.");
        }
        if (rawDiagnostics.Any(diagnostic => diagnostic != null &&
            (!diagnostic.IsRecoverable || diagnostic.Severity == OcrDiagnosticSeverity.Error))) {
            throw new InvalidDataException(
                "OCR reported an error or non-recoverable diagnostic, so recognition could not be accepted.");
        }
        if (rawSpans.Any(span => span != null && !string.IsNullOrEmpty(span.Text) &&
            !IsSupportedSpanLevel(span.Level))) {
            throw new InvalidDataException("OCR text spans must use line, word, or character granularity.");
        }

        OcrTextSpan[] spans = rawSpans
            .Select((span, index) => new { Span = span, Index = index })
            .Where(item => item.Span != null && IsSupportedSpanLevel(item.Span.Level) &&
                !string.IsNullOrEmpty(item.Span.Text) &&
                (!string.IsNullOrWhiteSpace(item.Span.Text) || item.Span.Level == OcrTextSpanLevel.Character))
            .OrderBy(item => item.Span.Sequence)
            .ThenBy(item => item.Span.Level)
            .ThenBy(item => item.Index)
            .Select(item => item.Span)
            .ToArray();
        if (spans.Length == 0 && !string.IsNullOrWhiteSpace(result.Text)) {
            throw new InvalidDataException(
                "OCR returned text without bounded line, word, or character geometry; concealment cannot be assessed.");
        }

        var builder = new OfficeContentSafetyBuilder(ReportFormat, options.Inspection);
        var targets = new Dictionary<string, RasterTarget>(StringComparer.Ordinal);
        var recognizedTargets = new List<RasterTarget>(spans.Length);
        var concealedTargets = new List<RasterConcealment>();
        int visibleSpans = 0;
        for (int index = 0; index < spans.Length; index++) {
            if ((index & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
            OcrTextSpan span = spans[index];
            ValidateConfidence(span.Confidence);
            PixelRegion region = ResolvePixelRegion(span, image.Width, image.Height);
            PixelRegion contrastRegion = Expand(region, 1, image.Width, image.Height);
            budget.ChargePixels(contrastRegion.Area);
            PixelEvidence pixels = InspectPixels(image, region, contrastRegion, cancellationToken);
            var target = new RasterTarget(span, region);
            recognizedTargets.Add(target);
            if (!TryClassify(region, pixels, options, out OfficeContentConcealmentKind kind, out string mechanism)) {
                visibleSpans++;
                continue;
            }

            bool canRedact = options.EnableOpaqueRectangleRedaction &&
                span.Confidence.HasValue &&
                span.Confidence.Value + 0.000001D >= options.MinimumOcrConfidenceForRedaction;
            OfficeContentCleanupCapability capability = canRedact
                ? OfficeContentCleanupCapability.RedactRegion
                : OfficeContentCleanupCapability.ReportOnly;
            concealedTargets.Add(new RasterConcealment(
                target,
                index,
                kind,
                mechanism,
                pixels,
                capability));
        }
        ValidateAggregateTextCoverage(
            result.Text,
            recognizedTargets,
            budget,
            cancellationToken);
        IReadOnlyDictionary<RasterTarget, IReadOnlyList<string>> instructionSignals =
            options.Inspection.DetectInstructionLikeText
                ? ResolveConcealedInstructionSignals(recognizedTargets, concealedTargets, cancellationToken)
                : new Dictionary<RasterTarget, IReadOnlyList<string>>();
        foreach (RasterConcealment concealed in concealedTargets) {
            cancellationToken.ThrowIfCancellationRequested();
            RasterTarget target = concealed.Target;
            if (string.IsNullOrWhiteSpace(target.Text)) continue;
            string location = "Frame[1]/Ocr[" + EngineLocationIdentity(engineId) + "]/" + target.Level +
                "[" + (concealed.Index + 1).ToString(CultureInfo.InvariantCulture) + "]@" +
                target.Region.Left.ToString(CultureInfo.InvariantCulture) + "," +
                target.Region.Top.ToString(CultureInfo.InvariantCulture) + "," +
                target.Region.Width.ToString(CultureInfo.InvariantCulture) + "," +
                target.Region.Height.ToString(CultureInfo.InvariantCulture);
            string evidence = concealed.Mechanism + " OCR provider '" + SanitizeIdentifier(engineId) + "' returned a bounded " +
                target.Level.ToString().ToLowerInvariant() + " region at " +
                target.Region.Left.ToString(CultureInfo.InvariantCulture) + "," +
                target.Region.Top.ToString(CultureInfo.InvariantCulture) + " with size " +
                target.Region.Width.ToString(CultureInfo.InvariantCulture) + "x" +
                target.Region.Height.ToString(CultureInfo.InvariantCulture) + " pixels. " +
                "The region plus its one-pixel perimeter has maximum pixel contrast " + concealed.Pixels.MaximumContrast.ToString("0.###", CultureInfo.InvariantCulture) +
                " and maximum alpha is " + concealed.Pixels.MaximumAlpha.ToString(CultureInfo.InvariantCulture) + "/255." +
                (target.Confidence.HasValue
                    ? " OCR confidence is " + target.Confidence.Value.ToString("0.###", CultureInfo.InvariantCulture) + "."
                    : " OCR confidence was not supplied, so cleanup remains report-only.");
            OfficeContentSafetyFinding finding = builder.AddWithInstructionSignals(
                concealed.Kind,
                OfficeContentSafetyRisk.ContextDependent,
                location,
                evidence,
                target.Text,
                instructionSignals.TryGetValue(target, out IReadOnlyList<string>? signals)
                    ? signals
                    : Array.Empty<string>(),
                concealed.Capability,
                inspectTextIntegrityEvidence: false);
            targets[finding.Id] = target;
        }

        foreach (OcrDiagnostic diagnostic in rawDiagnostics) {
            if (diagnostic == null) continue;
            builder.AddDiagnostic(
                "OCR provider '" + SanitizeIdentifier(engineId) + "' reported " +
                diagnostic.Severity.ToString().ToLowerInvariant() + " diagnostic '" +
                SanitizeIdentifier(diagnostic.Code) + "'. Provider messages and attributes are not treated as image text.");
        }
        builder.AddDiagnostic(
            "OCR provider '" + SanitizeIdentifier(engineId) + "' returned " + spans.Length.ToString(CultureInfo.InvariantCulture) +
            " bounded line, word, or character spans; " +
            visibleSpans.ToString(CultureInfo.InvariantCulture) + " had no bounded concealment evidence.");
        return new AnalysisState(image, builder.Build(), targets, recognizedTargets.AsReadOnly());
    }

    private static IReadOnlyDictionary<RasterTarget, IReadOnlyList<string>> ResolveConcealedInstructionSignals(
        IReadOnlyList<RasterTarget> targets,
        IReadOnlyList<RasterConcealment> concealments,
        CancellationToken cancellationToken) {
        var concealed = new HashSet<RasterTarget>(concealments.Select(item => item.Target));
        var signals = new Dictionary<RasterTarget, IReadOnlyList<string>>();
        AddLevelSignals(OcrTextSpanLevel.Line, " ");
        AddLevelSignals(OcrTextSpanLevel.Word, " ");
        AddLevelSignals(OcrTextSpanLevel.Character, string.Empty);
        AddMixedSignals();
        return signals;

        void AddLevelSignals(OcrTextSpanLevel level, string separator) {
            var run = new List<RasterTarget>();
            foreach (RasterTarget target in targets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (target.Level != level) continue;
                if (!concealed.Contains(target)) {
                    Flush(run, separator);
                    continue;
                }
                run.Add(target);
            }
            Flush(run, separator);
        }

        void AddMixedSignals() {
            var run = new List<RasterTarget>();
            foreach (RasterTarget target in targets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!concealed.Contains(target)) {
                    FlushMixed(run);
                    continue;
                }
                run.Add(target);
            }
            FlushMixed(run);
        }

        void Flush(List<RasterTarget> run, string separator) {
            if (run.Count == 0) return;
            string text = string.Join(separator, run.Select(item => item.Text));
            Register(run, OfficeContentInstructionDetector.Detect(text));
            run.Clear();
        }

        void FlushMixed(List<RasterTarget> run) {
            if (run.Count == 0) return;
            Register(run, OfficeContentInstructionDetector.Detect(
                FlattenTargetText(run, new HashSet<RasterTarget>(), cancellationToken)));
            run.Clear();
        }

        void Register(IReadOnlyList<RasterTarget> run, IReadOnlyList<string> detected) {
            if (detected.Count == 0) return;
            foreach (RasterTarget target in run) {
                if (string.IsNullOrWhiteSpace(target.Text)) continue;
                if (signals.TryGetValue(target, out IReadOnlyList<string>? existing)) {
                    signals[target] = existing.Concat(detected).Distinct(StringComparer.Ordinal).ToArray();
                } else {
                    signals[target] = detected;
                }
            }
        }
    }

    private static void ValidateAggregateTextCoverage(
        string? aggregateText,
        IReadOnlyList<RasterTarget> targets,
        RasterWorkBudget budget,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(aggregateText)) return;
        if (HasEquivalentTargetText(aggregateText, targets, OcrTextSpanLevel.Line, " ", cancellationToken) ||
            HasEquivalentTargetText(aggregateText, targets, OcrTextSpanLevel.Word, " ", cancellationToken) ||
            HasEquivalentTargetText(
                aggregateText,
                targets,
                OcrTextSpanLevel.Character,
                string.Empty,
                cancellationToken)) {
            return;
        }
        HashSet<RasterTarget> aggregateParents = ResolveAggregateParentTargets(
            targets,
            budget,
            cancellationToken);
        string flattened = FlattenTargetText(targets, aggregateParents, cancellationToken);
        if (string.Equals(
                NormalizeWhitespace(aggregateText, cancellationToken),
                NormalizeWhitespace(flattened, cancellationToken),
                StringComparison.Ordinal)) {
            return;
        }
        throw new InvalidDataException(
            "OCR aggregate text is not fully represented by accepted bounded text spans.");
    }

    private static bool HasEquivalentTargetText(
        string aggregateText,
        IReadOnlyList<RasterTarget> targets,
        OcrTextSpanLevel level,
        string separator,
        CancellationToken cancellationToken) {
        string[] text = targets
            .Where(target => target.Level == level)
            .OrderBy(target => target.Sequence)
            .Select(target => target.Text)
            .ToArray();
        return text.Length > 0 && string.Equals(
            NormalizeWhitespace(aggregateText, cancellationToken),
            NormalizeWhitespace(string.Join(separator, text), cancellationToken),
            StringComparison.Ordinal);
    }

    private static string FlattenTargetText(
        IReadOnlyList<RasterTarget> targets,
        IReadOnlySet<RasterTarget> aggregateParents,
        CancellationToken cancellationToken) {
        var flattened = new StringBuilder();
        RasterTarget? previous = null;
        for (int index = 0; index < targets.Count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            RasterTarget target = targets[index];
            if (aggregateParents.Contains(target)) continue;
            if (previous != null &&
                (previous.Level != OcrTextSpanLevel.Character || target.Level != OcrTextSpanLevel.Character)) {
                flattened.Append(' ');
            }
            flattened.Append(target.Text);
            previous = target;
        }
        return flattened.ToString();
    }

    private static string NormalizeWhitespace(string value, CancellationToken cancellationToken) {
        var normalized = new StringBuilder(value.Length);
        bool pendingSpace = false;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char character = value[index];
            if (char.IsWhiteSpace(character)) {
                pendingSpace = normalized.Length > 0;
                continue;
            }
            if (pendingSpace) normalized.Append(' ');
            normalized.Append(character);
            pendingSpace = false;
        }
        return normalized.ToString();
    }

    private static void ValidateOcrOutputCharacters(
        OcrResult result,
        IReadOnlyList<OcrTextSpan> spans,
        IReadOnlyList<OcrDiagnostic> diagnostics,
        OfficeRasterContentSafetyOptions.Snapshot options,
        CancellationToken cancellationToken) {
        long totalCharacters = 0L;
        int totalAttributes = 0;
        AddCharacters(result.Text);
        AddCharacters(result.Language);
        AddCharacters(result.Provider);
        AddCharacters(result.Model);
        AddCharacters(result.Orientation?.Script);
        foreach (OcrTextSpan? span in spans) {
            if (span == null) continue;
            AddCharacters(span.Text);
            AddCharacters(span.Language);
            AddCharacters(span.BlockId);
            AddCharacters(span.ParagraphId);
            AddCharacters(span.LineId);
        }
        foreach (OcrDiagnostic? diagnostic in diagnostics) {
            if (diagnostic == null) continue;
            AddCharacters(diagnostic.Code);
            AddCharacters(diagnostic.Message);
            AddCharacters(diagnostic.Source);
            IReadOnlyDictionary<string, string> attributes = diagnostic.Attributes ??
                new Dictionary<string, string>(StringComparer.Ordinal);
            totalAttributes = checked(totalAttributes + attributes.Count);
            if (totalAttributes > options.MaximumOcrSpans) {
                throw new InvalidDataException("The OCR result exceeds the configured diagnostic-attribute limit.");
            }
            foreach (KeyValuePair<string, string> attribute in attributes) {
                AddCharacters(attribute.Key);
                AddCharacters(attribute.Value);
            }
        }

        void AddCharacters(string? value) {
            totalCharacters = checked(totalCharacters + (value?.Length ?? 0));
            if (totalCharacters > options.Inspection.MaxCharacters) {
                throw new InvalidDataException("OCR output exceeds the configured character limit.");
            }
            if (value == null) return;
            for (int index = 0; index < value.Length; index++) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                char character = value[index];
                if (char.IsHighSurrogate(character)) {
                    if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1])) {
                        throw new InvalidDataException("OCR output contains malformed Unicode text.");
                    }
                    index++;
                } else if (char.IsLowSurrogate(character)) {
                    throw new InvalidDataException("OCR output contains malformed Unicode text.");
                }
            }
        }
    }

    private static bool IsSupportedSpanLevel(OcrTextSpanLevel level) =>
        level == OcrTextSpanLevel.Line || level == OcrTextSpanLevel.Word || level == OcrTextSpanLevel.Character;

    private static PixelRegion ResolvePixelRegion(OcrTextSpan span, int width, int height) {
        if (span.PageNumber.HasValue && span.PageNumber.Value != 1) {
            throw new InvalidDataException("A single-frame raster OCR span referenced another page.");
        }
        OcrRegion? source = span.Region;
        if (source == null) throw new InvalidDataException("OCR text spans require bounded geometry.");
        ValidateFinite(source.X, source.Y, source.Width, source.Height);
        if (source.Width <= 0D || source.Height <= 0D) {
            throw new InvalidDataException("OCR text spans require positive geometry.");
        }
        double x = source.X;
        double y = source.Y;
        double regionWidth = source.Width;
        double regionHeight = source.Height;
        if (span.CoordinateUnit == OcrCoordinateUnit.Normalized) {
            if (x < 0D || y < 0D || regionWidth > 1D || regionHeight > 1D ||
                x + regionWidth > 1D + 0.000001D || y + regionHeight > 1D + 0.000001D) {
                throw new InvalidDataException("Normalized OCR geometry must remain within zero through one.");
            }
            x *= width;
            regionWidth *= width;
            y *= height;
            regionHeight *= height;
        } else if (span.CoordinateUnit != OcrCoordinateUnit.Pixels) {
            throw new InvalidDataException("OCR text geometry must use pixels or normalized image coordinates.");
        }

        if (x < 0D || y < 0D || x + regionWidth > width + 0.000001D ||
            y + regionHeight > height + 0.000001D) {
            throw new InvalidDataException("OCR text geometry falls outside the decoded image.");
        }

        int left = checked((int)Math.Floor(x));
        int top = checked((int)Math.Floor(y));
        int right = checked((int)Math.Ceiling(x + regionWidth));
        int bottom = checked((int)Math.Ceiling(y + regionHeight));
        if (left < 0 || top < 0 || right > width || bottom > height || right <= left || bottom <= top) {
            throw new InvalidDataException("OCR text geometry falls outside the decoded image.");
        }
        return new PixelRegion(left, top, right, bottom);
    }

    private static PixelEvidence InspectPixels(
        OfficeRasterImage image,
        PixelRegion region,
        PixelRegion contrastRegion,
        CancellationToken cancellationToken) {
        byte[] pixels = image.PixelBuffer;
        double minimumBlack = 1D;
        double maximumBlack = 0D;
        double minimumWhite = 1D;
        double maximumWhite = 0D;
        byte maximumAlpha = 0;
        int cancellationCounter = 0;
        for (int y = contrastRegion.Top; y < contrastRegion.Bottom; y++) {
            int offset = ((y * image.Width) + contrastRegion.Left) * 4;
            for (int x = contrastRegion.Left; x < contrastRegion.Right; x++, offset += 4) {
                if ((cancellationCounter++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                byte alpha = pixels[offset + 3];
                if (x >= region.Left && x < region.Right && y >= region.Top && y < region.Bottom && alpha > maximumAlpha) {
                    maximumAlpha = alpha;
                }
                double normalizedAlpha = alpha / 255D;
                byte blackRed = ToByte(pixels[offset] * normalizedAlpha);
                byte blackGreen = ToByte(pixels[offset + 1] * normalizedAlpha);
                byte blackBlue = ToByte(pixels[offset + 2] * normalizedAlpha);
                byte whiteRed = ToByte(pixels[offset] * normalizedAlpha + 255D * (1D - normalizedAlpha));
                byte whiteGreen = ToByte(pixels[offset + 1] * normalizedAlpha + 255D * (1D - normalizedAlpha));
                byte whiteBlue = ToByte(pixels[offset + 2] * normalizedAlpha + 255D * (1D - normalizedAlpha));
                double black = OfficeColorContrast.RelativeLuminance(OfficeColor.FromRgb(blackRed, blackGreen, blackBlue));
                double white = OfficeColorContrast.RelativeLuminance(OfficeColor.FromRgb(whiteRed, whiteGreen, whiteBlue));
                minimumBlack = Math.Min(minimumBlack, black);
                maximumBlack = Math.Max(maximumBlack, black);
                minimumWhite = Math.Min(minimumWhite, white);
                maximumWhite = Math.Max(maximumWhite, white);
            }
        }
        double blackContrast = (maximumBlack + 0.05D) / (minimumBlack + 0.05D);
        double whiteContrast = (maximumWhite + 0.05D) / (minimumWhite + 0.05D);
        return new PixelEvidence(Math.Max(blackContrast, whiteContrast), maximumAlpha);
    }

    private static bool TryClassify(
        PixelRegion region,
        PixelEvidence pixels,
        OfficeRasterContentSafetyOptions.Snapshot options,
        out OfficeContentConcealmentKind kind,
        out string evidence) {
        if (pixels.MaximumAlpha <= options.MaximumConcealedAlpha) {
            kind = OfficeContentConcealmentKind.TransparentText;
            evidence = "Every pixel in the OCR region is below the configured alpha ceiling.";
            return true;
        }
        if (region.Height <= options.MaximumTinyTextHeightPixels) {
            kind = OfficeContentConcealmentKind.TinyText;
            evidence = "The bounded OCR region is at or below the configured tiny-text pixel height.";
            return true;
        }
        if (pixels.MaximumContrast + 0.000001D < options.Inspection.MinimumVisibleContrastRatio) {
            kind = OfficeContentConcealmentKind.LowContrastText;
            evidence = "Even the maximum pairwise pixel contrast across the OCR region and its one-pixel perimeter is below the configured visibility threshold.";
            return true;
        }
        kind = default;
        evidence = string.Empty;
        return false;
    }

    private static void ValidateConfidence(double? confidence) {
        if (!confidence.HasValue) return;
        if (double.IsNaN(confidence.Value) || double.IsInfinity(confidence.Value) ||
            confidence.Value < 0D || confidence.Value > 1D) {
            throw new InvalidDataException("OCR confidence must be between zero and one.");
        }
    }

    private static void ValidateFinite(params double[] values) {
        if (values.Any(value => double.IsNaN(value) || double.IsInfinity(value))) {
            throw new InvalidDataException("OCR geometry must contain only finite values.");
        }
    }

    private static string SanitizeIdentifier(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return "unspecified";
        var characters = value.Trim().Take(128)
            .Select(character => char.IsLetterOrDigit(character) || character is '-' or '_' or '.' ? character : '_')
            .ToArray();
        return characters.Length == 0 ? "unspecified" : new string(characters);
    }

    private static string EngineLocationIdentity(string engineId) =>
        SanitizeIdentifier(engineId) + "-" + HashText(engineId).Substring(0, 12);

    private static string HashText(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static byte ToByte(double value) => (byte)Math.Min(255D, Math.Max(0D, Math.Round(value)));

    private readonly struct PixelEvidence {
        internal PixelEvidence(double maximumContrast, byte maximumAlpha) {
            MaximumContrast = maximumContrast;
            MaximumAlpha = maximumAlpha;
        }

        internal double MaximumContrast { get; }
        internal byte MaximumAlpha { get; }
    }

    private sealed class RasterConcealment {
        internal RasterConcealment(
            RasterTarget target,
            int index,
            OfficeContentConcealmentKind kind,
            string mechanism,
            PixelEvidence pixels,
            OfficeContentCleanupCapability capability) {
            Target = target;
            Index = index;
            Kind = kind;
            Mechanism = mechanism;
            Pixels = pixels;
            Capability = capability;
        }

        internal RasterTarget Target { get; }
        internal int Index { get; }
        internal OfficeContentConcealmentKind Kind { get; }
        internal string Mechanism { get; }
        internal PixelEvidence Pixels { get; }
        internal OfficeContentCleanupCapability Capability { get; }
    }
}
