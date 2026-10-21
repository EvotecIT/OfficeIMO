using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private const int DiagnosticAlphaPixelLimit = 1024 * 1024;
        private const int DiagnosticAlphaRunLimit = 65536;
        private readonly Dictionary<PdfImageStream, byte[]> _diagnosticAlpha = new();

        /// <summary>Projects the writer's own image placement and alpha samples, retaining its intrinsic crop before assessing external loss.</summary>
        private IEnumerable<(List<List<OfficePoint>> Contours, List<OfficeTextInkClip> Clips)> DiagnosticInlineImagePaint(
            PdfInlineImage image, double x, double baseline) {
            ImageBlock block = image.Block;
            double bottom = baseline + image.BaselineOffset;
            PageImage placed = CreatePageImage(block, block.Style ?? new PdfImageStyle(),
                image.TableTextRotation == 0 ? x : x + (image.Width - image.Height) / 2D,
                image.TableTextRotation == 0 ? bottom : bottom + (image.Height - image.Width) / 2D,
                image.TableTextRotation == 0 ? image.Width : image.Height,
                image.TableTextRotation == 0 ? image.Height : image.Width);
            if (image.TableTextRotation != 0) {
                placed.RotationAngle -= image.TableTextRotation;
                placed.RotationCenterX = x + image.Width / 2D;
                placed.RotationCenterY = bottom + image.Height / 2D;
            }
            PdfImageStream stream;
            if (block.PreparedStream != null) stream = block.PreparedStream;
            else if (!TryBuildImageStream(block.Data, block.Info, block.Width, block.Height, cancellationToken, out stream, out _))
                throw new NotSupportedException("The inline image stream could not be inspected.");
            OfficeTransform placement = new OfficeImageProjection(new OfficeImagePlacement(placed.X, placed.Y, placed.W, placed.H),
                rotationDegrees: placed.RotationAngle, rotationCenterX: placed.RotationCenterX, rotationCenterY: placed.RotationCenterY,
                flipHorizontal: placed.HorizontalFlip, flipVertical: placed.VerticalFlip).CreateUnitSquareTransform();
            OfficeTransform sampleToLocal = GetImageOrientationTransform(stream.Orientation).Then(placement);
            var clips = new List<OfficeTextInkClip>();
            if (placed.SourceCrop?.HasCrop == true) {
                PdfImageSourceCrop crop = placed.SourceCrop;
                clips.Add(DiagnosticRectangleClip(new TableCellContentFrame(crop.Left, 1D - crop.Top,
                    1D - crop.Left - crop.Right, 1D - crop.Top - crop.Bottom), placement));
            }
            if (placed.ClipPath != null) {
                OfficeTransform clipPlacement = new(1D, 0D, 0D, -1D, placed.ClipX, placed.ClipY + placed.ClipHeight);
                var rings = OfficeClipPathGeometry.CreateContours(placed.ClipPath, points => {
                    var mapped = new List<OfficePoint>(points.Count);
                    foreach (var point in points) { cancellationToken.ThrowIfCancellationRequested(); mapped.Add(clipPlacement.TransformPoint(point)); }
                    return mapped;
                });
                clips.Add(OfficeTextInkClip.FromFilledContours(rings, placed.ClipPath.FillRule));
            }
            if (stream.SoftMask == null) {
                yield return (MapRectangle(0D, 0D, 1D, 1D), clips);
                yield break;
            }
            PdfImageStream mask = stream.SoftMask;
            byte[] alpha = DiagnosticImageAlpha(mask);
            int sampleBytes = mask.BitsPerComponent / 8;
            int runCount = 0;
            for (int row = 0; row < mask.PixelHeight; row++) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int column = 0; column < mask.PixelWidth;) {
                    while (column < mask.PixelWidth && !HasAlpha(column, row)) column++;
                    int start = column;
                    while (column < mask.PixelWidth && HasAlpha(column, row)) column++;
                    if (start == column) continue;
                    if (++runCount > DiagnosticAlphaRunLimit) throw new NotSupportedException("Inline image alpha exceeds the bounded run inspection limit.");
                    yield return (MapRectangle((double)start / mask.PixelWidth, 1D - (double)(row + 1) / mask.PixelHeight,
                        (double)(column - start) / mask.PixelWidth, 1D / mask.PixelHeight), clips);
                }
            }

            bool HasAlpha(int column, int row) {
                if ((column & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                int offset = (row * mask.PixelWidth + column) * sampleBytes;
                int value = sampleBytes == 1 ? alpha[offset] : ReadUInt16BigEndian(alpha, offset);
                // Match the existing writer's reduction for a pre-1.5 target, including alpha.
                return sampleBytes == 1 || currentOpts.FileVersion >= PdfFileVersion.Pdf15 ? value > 0 : Scale16BitSampleToByte(value) > 0;
            }
            List<List<OfficePoint>> MapRectangle(double left, double bottomY, double width, double height) {
                var contours = DiagnosticRectangle(left, bottomY, width, height);
                foreach (var ring in contours) for (int index = 0; index < ring.Count; index++) ring[index] = sampleToLocal.TransformPoint(ring[index]);
                return contours;
            }
        }

        /// <summary>Checks nonempty raster paint in line slots discarded before the ordinary row renderer is entered.</summary>
        private void ReportOmittedInlineImagePaint(TableCellLayout cell, TableCellTextLayout layout, int row, int start, int end,
            double x, double top, double width, double height) {
            string source = "PdfTableCell[" + row + "," + cell.Column + "]";
            try {
                for (int line = start; line < Math.Min(end, layout.Lines.Count); line++) {
                    foreach (RichSeg segment in layout.Lines[line]) {
                        if (segment.InlineElement is not PdfInlineImage image) continue;
                        foreach (var paint in DiagnosticInlineImagePaint(image, 0D, 0D)) {
                            var bounds = OfficeRasterCanvas.MeasureNominalFilledContourBounds(paint.Contours, OfficeFillRule.NonZero, paint.Clips, cancellationToken);
                            RequireMeasured(bounds.IsMeasured);
                            if (!bounds.HasInk) continue;
                            if (_diagnosticCanvasClips.Count > 0 || !_canvasEffectToPage.Equals(OfficeTransform.Identity))
                                currentOpts.AddUnassessedLayoutWarning("TableCellPaintBoundsUnmeasured", source,
                                    "An omitted raster line has nonzero intrinsic paint; its interaction with the enclosing effect or clip cannot be measured after admission.");
                            else ReportAuthoredLayoutDiagnostic("TableCellInlinePaintOmitted", source,
                                "A raster inline with nonzero visible alpha was omitted with its line when it exceeded the authored cell height.",
                                PdfLayoutDiagnosticKind.ClippedContent, x, top, width, height);
                            return;
                        }
                    }
                }
            } catch (NotSupportedException exception) {
                currentOpts.AddUnassessedLayoutWarning("TableCellPaintBoundsUnmeasured", source, exception.Message);
            }
        }

        /// <summary>Reads prepared alpha through the existing bounded filter/predictor owners; no color decode or output descriptor is mutated.</summary>
        private byte[] DiagnosticImageAlpha(PdfImageStream mask) {
            if (_diagnosticAlpha.TryGetValue(mask, out byte[]? cached)) return cached;
            long pixels = (long)mask.PixelWidth * mask.PixelHeight;
            if (mask.PixelWidth <= 0 || mask.PixelHeight <= 0 || pixels > DiagnosticAlphaPixelLimit ||
                mask.BitsPerComponent is not (8 or 16) || mask.Data.Length > 4 * DiagnosticAlphaPixelLimit ||
                mask.DictionarySuffix != BuildPngPredictorDictionarySuffix("/DeviceGray", 1, mask.PixelWidth, mask.BitsPerComponent))
                throw new NotSupportedException("Inline image alpha exceeds the bounded sample inspection contract.");
            int bytes = checked((int)pixels * (mask.BitsPerComponent / 8));
            if (!FlateDecoder.TryDecode(mask.Data, checked(bytes + mask.PixelHeight), out byte[] filtered, out _, cancellationToken))
                throw new NotSupportedException("The inline image alpha filter could not be inspected within its budget.");
            byte[] decoded;
            try {
                decoded = PngPredictorDecoder.Decode(filtered, mask.PixelWidth, 1, mask.BitsPerComponent, bytes, cancellationToken);
            } catch (Exception exception) when (exception is InvalidDataException or FormatException) {
                throw new NotSupportedException("The inline image alpha predictor could not be inspected.", exception);
            }
            if (decoded.Length != bytes) throw new NotSupportedException("The prepared inline image alpha has incomplete samples.");
            if (_diagnosticAlpha.Count >= 4) _diagnosticAlpha.Clear();
            _diagnosticAlpha.Add(mask, decoded);
            return decoded;
        }
    }
}
