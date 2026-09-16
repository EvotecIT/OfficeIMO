using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;

namespace OfficeIMO.Workflows;

public static partial class OfficeRasterContentSafety {
    /// <summary>
    /// Reinspects an exact raster snapshot, covers only selected cleanup-capable OCR regions with the configured
    /// opaque color, emits a single-frame PNG derivative, reopens it, and reruns OCR-backed inspection.
    /// </summary>
    public static async Task<OfficeContentCleanupResult> RedactSelectedContentAsync(
        byte[] imageBytes,
        IOcrEngine engine,
        OfficeContentCleanupSelection selection,
        OfficeRasterContentSafetyOptions options,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(imageBytes);
        ArgumentNullException.ThrowIfNull(engine);
        ArgumentNullException.ThrowIfNull(selection);
        ArgumentNullException.ThrowIfNull(options);
        OfficeRasterContentSafetyOptions.Snapshot snapshot = options.Capture();
        if (!snapshot.EnableOpaqueRectangleRedaction) {
            throw new InvalidOperationException(
                "Opaque rectangle redaction must be explicitly enabled before raster content can be changed.");
        }

        OfficeContentSafetyInputGuard.ValidateBytes(imageBytes, snapshot.Inspection);
        byte[] input = (byte[])imageBytes.Clone();
        OcrEngineExecution execution = OcrEngineRunner.CreateExecution(engine);
        AnalysisState beforeState = await InspectCoreAsync(input, execution, snapshot, cancellationToken)
            .ConfigureAwait(false);
        IReadOnlyList<OfficeContentSafetyFinding> selected =
            OfficeContentSafetyBuilder.ResolveSelection(beforeState.Report, selection);
        if (selected.Any(item => item.CleanupCapability != OfficeContentCleanupCapability.RedactRegion)) {
            throw new InvalidOperationException("Raster cleanup accepts only bounded region-redaction findings.");
        }
        if (selected.Count == 0) {
            return new OfficeContentCleanupResult(
                input,
                beforeState.Report,
                beforeState.Report,
                Array.Empty<OfficeContentCleanupChange>());
        }

        var selectedTargets = new List<RasterTarget>(selected.Count);
        foreach (OfficeContentSafetyFinding finding in selected) {
            if (!beforeState.Targets.TryGetValue(finding.Id, out RasterTarget? target)) {
                throw new InvalidDataException("A selected raster finding no longer has exact region evidence.");
            }
            selectedTargets.Add(target);
        }

        var changedRegions = new List<PixelRegion>(selectedTargets.Count);
        var selectedTargetSet = new HashSet<RasterTarget>(selectedTargets);
        RasterTarget[] unselectedTargets = beforeState.RecognizedTargets
            .Where(target => !selectedTargetSet.Contains(target))
            .ToArray();
        long remainingRedactionWork = snapshot.MaximumPixelAnalysisWork;
        long remainingRegionComparisons = snapshot.MaximumRegionComparisons;
        int regionComparisonCount = 0;
        for (int index = 0; index < selectedTargets.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PixelRegion expanded = Expand(
                selectedTargets[index].Region,
                snapshot.RedactionPaddingPixels,
                beforeState.Image.Width,
                beforeState.Image.Height);
            foreach (RasterTarget target in unselectedTargets) {
                ChargeRegionComparison(
                    ref remainingRegionComparisons,
                    ref regionComparisonCount,
                    cancellationToken);
                if (expanded.Intersects(target.Region)) {
                    throw new InvalidOperationException(
                        "A selected raster redaction region overlaps recognized text that was not selected.");
                }
            }
            long regionWork = checked(expanded.Area * 2L);
            if (regionWork > remainingRedactionWork) {
                throw new InvalidDataException("Raster redaction work exceeds the configured pixel-analysis limit.");
            }
            remainingRedactionWork -= regionWork;
            changedRegions.Add(expanded);
        }

        OfficeRasterImage redacted = OfficeRasterImage.FromRgba32(
            beforeState.Image.Width,
            beforeState.Image.Height,
            beforeState.Image.PixelBuffer);
        foreach (PixelRegion region in changedRegions) {
            cancellationToken.ThrowIfCancellationRequested();
            Fill(redacted, region, snapshot.RedactionColor, cancellationToken);
        }

        byte[] output = OfficeRasterImageEncoder.Encode(
            redacted,
            OfficeImageExportFormat.Png,
            options: null,
            snapshot.MaximumOutputBytes,
            cancellationToken);
        VerifyRedactionOutput(output, redacted, changedRegions, snapshot, cancellationToken);

        AnalysisState afterState = await InspectCoreAsync(output, execution, snapshot, cancellationToken)
            .ConfigureAwait(false);
        for (int selectedIndex = 0; selectedIndex < selectedTargets.Count; selectedIndex++) {
            PixelRegion changedRegion = changedRegions[selectedIndex];
            foreach (RasterTarget afterTarget in afterState.RecognizedTargets) {
                ChargeRegionComparison(
                    ref remainingRegionComparisons,
                    ref regionComparisonCount,
                    cancellationToken);
                if (!changedRegion.Intersects(afterTarget.Region)) continue;
                throw new InvalidDataException(
                    "OCR still recognized text inside a changed redaction region; output was not accepted.");
            }
        }

        OfficeContentCleanupChange[] changes = selected
            .Select(item => new OfficeContentCleanupChange(
                item.Id,
                item.Location,
                OfficeContentCleanupCapability.RedactRegion))
            .ToArray();
        return new OfficeContentCleanupResult(output, beforeState.Report, afterState.Report, changes);
    }

    private static void ChargeRegionComparison(
        ref long remaining,
        ref int comparisonCount,
        CancellationToken cancellationToken) {
        if (remaining <= 0L) {
            throw new InvalidDataException("Raster redaction geometry comparisons exceed the configured limit.");
        }
        remaining--;
        if ((comparisonCount++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
    }

    private static PixelRegion Expand(PixelRegion region, int padding, int width, int height) =>
        new PixelRegion(
            Math.Max(0, region.Left - padding),
            Math.Max(0, region.Top - padding),
            Math.Min(width, region.Right + padding),
            Math.Min(height, region.Bottom + padding));

    private static void Fill(
        OfficeRasterImage image,
        PixelRegion region,
        OfficeColor color,
        CancellationToken cancellationToken) {
        byte[] pixels = image.PixelBuffer;
        int cancellationCounter = 0;
        for (int y = region.Top; y < region.Bottom; y++) {
            int offset = ((y * image.Width) + region.Left) * 4;
            for (int x = region.Left; x < region.Right; x++, offset += 4) {
                if ((cancellationCounter++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                pixels[offset] = color.R;
                pixels[offset + 1] = color.G;
                pixels[offset + 2] = color.B;
                pixels[offset + 3] = byte.MaxValue;
            }
        }
    }

    private static void VerifyRedactionOutput(
        byte[] output,
        OfficeRasterImage expected,
        IReadOnlyList<PixelRegion> changedRegions,
        OfficeRasterContentSafetyOptions.Snapshot options,
        CancellationToken cancellationToken) {
        var decodeOptions = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = checked((int)Math.Min(options.MaximumOutputBytes, int.MaxValue)),
            MaximumDecodedPixels = options.MaximumDecodedPixels,
            FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames,
            CancellationToken = cancellationToken
        };
        if (!OfficeRasterImageDecoder.TryDecode(
                output,
                decodeOptions,
                out OfficeRasterImage? reopened,
                out OfficeRasterDecodeInfo info) || reopened == null) {
            throw new InvalidDataException(info.Diagnostic ?? "The redacted PNG could not be reopened.");
        }
        if (reopened.Width != expected.Width || reopened.Height != expected.Height) {
            throw new InvalidDataException("The redacted PNG changed the source dimensions.");
        }
        byte[] pixels = reopened.PixelBuffer;
        if (!pixels.AsSpan().SequenceEqual(expected.PixelBuffer)) {
            throw new InvalidDataException("The reopened PNG did not preserve the exact normalized redaction pixels.");
        }
        foreach (PixelRegion region in changedRegions) {
            int cancellationCounter = 0;
            for (int y = region.Top; y < region.Bottom; y++) {
                int offset = ((y * reopened.Width) + region.Left) * 4;
                for (int x = region.Left; x < region.Right; x++, offset += 4) {
                    if ((cancellationCounter++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (pixels[offset] != options.RedactionColor.R ||
                        pixels[offset + 1] != options.RedactionColor.G ||
                        pixels[offset + 2] != options.RedactionColor.B ||
                        pixels[offset + 3] != byte.MaxValue) {
                        throw new InvalidDataException("The reopened PNG did not preserve an opaque selected redaction region.");
                    }
                }
            }
        }
    }
}