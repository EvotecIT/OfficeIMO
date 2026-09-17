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
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterContentSafetyOptions.Snapshot snapshot = options.Capture();
        if (!snapshot.EnableOpaqueRectangleRedaction) {
            throw new InvalidOperationException(
                "Opaque rectangle redaction must be explicitly enabled before raster content can be changed.");
        }

        OfficeContentSafetyInputGuard.ValidateBytes(imageBytes, snapshot.Inspection);
        byte[] input = (byte[])imageBytes.Clone();
        OcrEngineExecution execution = OcrEngineRunner.CreateExecution(engine);
        var budget = new RasterWorkBudget(snapshot);
        long callerRetainedBytes = imageBytes.LongLength + 24L;
        AnalysisState beforeState = await InspectCoreAsync(
                input,
                execution,
                snapshot,
                budget,
                callerRetainedBytes,
                cancellationToken)
            .ConfigureAwait(false);
        IReadOnlyList<OfficeContentSafetyFinding> selected =
            OfficeContentSafetyBuilder.ResolveSelection(beforeState.Report, selection);
        if (selected.Any(item => item.CleanupCapability != OfficeContentCleanupCapability.RedactRegion)) {
            throw new InvalidOperationException("Raster cleanup accepts only bounded region-redaction findings.");
        }
        if (selected.Count == 0) {
            return OfficeContentCleanupResult.FromOwnedOutput(
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
        HashSet<RasterTarget> beforeAggregateParents = ResolveAggregateParentTargets(
            beforeState.RecognizedTargets,
            budget,
            cancellationToken);
        for (int index = 0; index < selectedTargets.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PixelRegion expanded = Expand(
                selectedTargets[index].Region,
                snapshot.RedactionPaddingPixels,
                beforeState.Image.Width,
                beforeState.Image.Height);
            foreach (RasterTarget target in unselectedTargets) {
                budget.ChargeComparison(cancellationToken);
                if (beforeAggregateParents.Contains(target)) continue;
                if (expanded.Intersects(target.Region)) {
                    throw new InvalidOperationException(
                        "A selected raster redaction region overlaps recognized text that was not selected.");
                }
            }
            budget.ChargePixels(checked(expanded.Area * 3L));
            if (!WouldChange(
                    beforeState.Image,
                    selectedTargets[index].Region,
                    snapshot.RedactionColor,
                    cancellationToken)) {
                throw new InvalidOperationException(
                    "A selected raster region already matches the configured redaction color and cannot produce a verified cleanup change.");
            }
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
            CreateMetadataFreePngEncodingOptions(),
            snapshot.MaximumOutputBytes,
            cancellationToken,
            checked(input.LongLength + 24L + callerRetainedBytes + beforeState.Image.PixelBuffer.LongLength + 24L));
        long retainedForOutputDecode = checked(
            input.LongLength + 24L + callerRetainedBytes +
            beforeState.Image.PixelBuffer.LongLength + 24L +
            redacted.PixelBuffer.LongLength + 24L);
        VerifyRedactionOutput(
            output,
            redacted,
            changedRegions,
            snapshot,
            budget,
            retainedForOutputDecode,
            cancellationToken);

        AnalysisState afterState = await InspectCoreAsync(
                output,
                execution,
                snapshot,
                budget,
                retainedForOutputDecode,
                cancellationToken)
            .ConfigureAwait(false);
        HashSet<RasterTarget> afterAggregateParents = ResolveAggregateParentTargets(
            afterState.RecognizedTargets,
            budget,
            cancellationToken);
        for (int selectedIndex = 0; selectedIndex < selectedTargets.Count; selectedIndex++) {
            PixelRegion changedRegion = changedRegions[selectedIndex];
            foreach (RasterTarget afterTarget in afterState.RecognizedTargets) {
                budget.ChargeComparison(cancellationToken);
                if (afterAggregateParents.Contains(afterTarget)) continue;
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
        return OfficeContentCleanupResult.FromOwnedOutput(output, beforeState.Report, afterState.Report, changes);
    }

    private static Dictionary<string, IReadOnlyList<RasterTarget>> IndexFinerTargetsByLine(
        IReadOnlyList<RasterTarget> targets) {
        var mutable = new Dictionary<string, List<RasterTarget>>(StringComparer.Ordinal);
        foreach (RasterTarget target in targets) {
            if (target.Level == OcrTextSpanLevel.Line || string.IsNullOrWhiteSpace(target.LineId)) continue;
            if (!mutable.TryGetValue(target.LineId!, out List<RasterTarget>? children)) {
                children = new List<RasterTarget>();
                mutable[target.LineId!] = children;
            }
            children.Add(target);
        }
        return mutable.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<RasterTarget>)pair.Value.AsReadOnly(),
            StringComparer.Ordinal);
    }

    private static HashSet<RasterTarget> ResolveAggregateParentTargets(
        IReadOnlyList<RasterTarget> targets,
        RasterWorkBudget budget,
        CancellationToken cancellationToken) {
        Dictionary<string, IReadOnlyList<RasterTarget>> childrenByLine = IndexFinerTargetsByLine(targets);
        var aggregates = new HashSet<RasterTarget>();
        foreach (RasterTarget candidate in targets) {
            if ((candidate.Level != OcrTextSpanLevel.Line && candidate.Level != OcrTextSpanLevel.Word) ||
                string.IsNullOrWhiteSpace(candidate.LineId) ||
                !childrenByLine.TryGetValue(candidate.LineId!, out IReadOnlyList<RasterTarget>? children)) {
                continue;
            }
            if (IsParentFullyRepresentedByChildren(
                    candidate,
                    children,
                    budget,
                    cancellationToken)) {
                aggregates.Add(candidate);
            }
        }
        return aggregates;
    }

    private static bool IsParentFullyRepresentedByChildren(
        RasterTarget parent,
        IReadOnlyList<RasterTarget> children,
        RasterWorkBudget budget,
        CancellationToken cancellationToken) {
        var words = new List<RasterTarget>();
        var characters = new List<RasterTarget>();
        foreach (RasterTarget child in children) {
            budget.ChargeComparison(cancellationToken);
            if (!parent.Region.Contains(child.Region) || child.Level <= parent.Level) continue;
            if (child.Level == OcrTextSpanLevel.Word) words.Add(child);
            else if (child.Level == OcrTextSpanLevel.Character) characters.Add(child);
        }
        return (parent.Level == OcrTextSpanLevel.Line &&
                HasEquivalentText(parent.Text, words, " ", cancellationToken)) ||
            HasEquivalentText(parent.Text, characters, string.Empty, cancellationToken);
    }

    private static bool HasEquivalentText(
        string lineText,
        IReadOnlyList<RasterTarget> children,
        string separator,
        CancellationToken cancellationToken) {
        if (children.Count == 0) return false;
        string childText = string.Join(
            separator,
            children.OrderBy(child => child.Sequence).Select(child => child.Text));
        return string.Equals(
            NormalizeWhitespace(lineText, cancellationToken),
            NormalizeWhitespace(childText, cancellationToken),
            StringComparison.Ordinal);
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

    private static bool WouldChange(
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
                if (pixels[offset] != color.R || pixels[offset + 1] != color.G ||
                    pixels[offset + 2] != color.B || pixels[offset + 3] != byte.MaxValue) {
                    return true;
                }
            }
        }
        return false;
    }

    private static void VerifyRedactionOutput(
        byte[] output,
        OfficeRasterImage expected,
        IReadOnlyList<PixelRegion> changedRegions,
        OfficeRasterContentSafetyOptions.Snapshot options,
        RasterWorkBudget budget,
        long additionalRetainedManagedBytes,
        CancellationToken cancellationToken) {
        var decodeOptions = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = checked((int)Math.Min(options.MaximumOutputBytes, int.MaxValue)),
            MaximumDecodedPixels = options.MaximumDecodedPixels,
            FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames,
            CancellationToken = cancellationToken,
            RetainedManagedBytes = additionalRetainedManagedBytes
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
        budget.ChargePixels(checked((long)reopened.Width * reopened.Height));
        if (!PixelBuffersEqual(pixels, expected.PixelBuffer, cancellationToken)) {
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

    internal static bool PixelBuffersEqual(
        byte[] actual,
        byte[] expected,
        CancellationToken cancellationToken) {
        if (actual.Length != expected.Length) return false;
        const int chunkSize = 64 * 1024;
        for (int offset = 0; offset < actual.Length; offset += chunkSize) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, actual.Length - offset);
            if (!actual.AsSpan(offset, count).SequenceEqual(expected.AsSpan(offset, count))) return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }
}
