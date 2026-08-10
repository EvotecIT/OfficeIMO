using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Canonical placement discovery and mutation owner for images on existing PDF pages.</summary>
internal static class PdfImageEditor {
    private const double CoordinateTolerance = 0.01D;
    private const double TransformTolerance = 0.0001D;

    internal static IReadOnlyList<PdfImagePlacement> Placements(byte[] pdf, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfImageExtractor.ExtractImagePlacements(PdfReadDocument.Open(pdf, readOptions));
    }

    internal static IReadOnlyList<PdfImagePlacement> Find(byte[] pdf, PdfPageRegion region, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        return document.Pages[region.PageNumber - 1]
            .GetImagePlacements(region.PageNumber)
            .Where(placement => Intersects(region, placement))
            .ToArray();
    }

    internal static ImageMutationResult Add(byte[] pdf, PdfPageRegion target, byte[] imageBytes, PdfImageEditOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(target, nameof(target));
        Guard.NotNull(imageBytes, nameof(imageBytes));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        ValidatePage(target.PageNumber, document.Pages.Count, nameof(target));
        PdfImageEditOptions snapshot = (options ?? new PdfImageEditOptions()).Snapshot();
        byte[] output = PdfStamper.StampImage(pdf, imageBytes, CreateStampOptions(target.PageNumber, target.X, target.Y, target.Width, target.Height, 0D, snapshot), readOptions);
        return new ImageMutationResult(output, 1);
    }

    internal static ImageMutationResult Remove(byte[] pdf, PdfImagePlacement placement, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfImagePlacement current = ResolveUniquePlacement(pdf, placement, readOptions);
        EnsureRemovablePlacement(current);
        byte[] output = PdfRedactionApplier.RemoveImagePlacements(pdf, new[] { current }, readOptions);
        EnsurePlacementRemoved(output, current, PdfReadOptions.WithMinimumInputBytes(readOptions, output.LongLength));
        return new ImageMutationResult(output, 1);
    }

    internal static ImageMutationResult Replace(byte[] pdf, PdfImagePlacement placement, byte[] imageBytes, PdfImageEditOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(imageBytes, nameof(imageBytes));
        PdfImagePlacement current = ResolveUniquePlacement(pdf, placement, readOptions);
        EnsureRemovablePlacement(current);
        ImageTransform transform = ResolvePortableTransform(current);
        PdfImageEditOptions snapshot = (options ?? new PdfImageEditOptions()).Snapshot();
        byte[] removed = PdfRedactionApplier.RemoveImagePlacements(pdf, new[] { current }, readOptions);
        PdfReadOptions afterRemovalOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, removed.LongLength);
        EnsurePlacementRemoved(removed, current, afterRemovalOptions);
        byte[] output = PdfStamper.StampImage(
            removed,
            imageBytes,
            CreateStampOptions(current.PageNumber, current.E, current.F, transform.Width, transform.Height, transform.RotationDegrees, snapshot),
            afterRemovalOptions);
        return new ImageMutationResult(output, 1);
    }

    internal static ImageMutationResult Move(byte[] pdf, PdfImagePlacement placement, double deltaX, double deltaY, PdfImageEditOptions? options, PdfReadOptions? readOptions) {
        ValidateFinite(deltaX, nameof(deltaX));
        ValidateFinite(deltaY, nameof(deltaY));
        Guard.NotNull(pdf, nameof(pdf));
        PdfImagePlacement current = ResolveUniquePlacement(pdf, placement, readOptions);
        EnsureRemovablePlacement(current);
        ImageTransform transform = ResolvePortableTransform(current);
        PdfExtractedImage image = ResolveMovableImage(pdf, current, readOptions);
        PdfImageEditOptions snapshot = (options ?? new PdfImageEditOptions()).Snapshot();
        byte[] removed = PdfRedactionApplier.RemoveImagePlacements(pdf, new[] { current }, readOptions);
        PdfReadOptions afterRemovalOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, removed.LongLength);
        EnsurePlacementRemoved(removed, current, afterRemovalOptions);
        byte[] output = PdfStamper.StampImage(
            removed,
            image.Bytes,
            CreateStampOptions(current.PageNumber, current.E + deltaX, current.F + deltaY, transform.Width, transform.Height, transform.RotationDegrees, snapshot),
            afterRemovalOptions);
        return new ImageMutationResult(output, 1);
    }

    private static PdfImagePlacement ResolveUniquePlacement(byte[] pdf, PdfImagePlacement placement, PdfReadOptions? readOptions) {
        Guard.NotNull(placement, nameof(placement));
        PdfImagePlacement[] matches = Placements(pdf, readOptions)
            .Where(candidate => SamePlacementIdentity(candidate, placement))
            .ToArray();
        if (matches.Length == 0) {
            throw new InvalidOperationException("The selected image placement does not exist in the current PDF document.");
        }
        if (matches.Length > 1) {
            throw new InvalidOperationException("The selected image placement is ambiguous because multiple invocations have the same resource, transform, and bounds. Select a uniquely placed image before editing.");
        }
        return matches[0];
    }

    private static PdfExtractedImage ResolveMovableImage(byte[] pdf, PdfImagePlacement placement, PdfReadOptions? readOptions) {
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        PdfImagePlacement[] currentPlacements = document.Pages[placement.PageNumber - 1]
            .GetImagePlacements(placement.PageNumber)
            .Where(candidate => SamePlacementIdentity(candidate, placement))
            .ToArray();
        if (currentPlacements.Length != 1) {
            throw new NotSupportedException("The selected image placement could not be resolved uniquely for a safe move.");
        }
        PdfImagePlacement current = currentPlacements[0];
        IReadOnlyList<PdfExtractedImage> images = document.Pages[placement.PageNumber - 1]
            .GetImages(placement.PageNumber, new[] { current });
        PdfExtractedImage[] matches = images.Where(image =>
            image.PageNumber == current.PageNumber &&
            string.Equals(image.ResourceName, current.ResourceName, StringComparison.Ordinal) &&
            image.ObjectNumber == current.ObjectNumber &&
            (image.ObjectNumber > 0 || image.DirectStreamIdentity == current.DirectStreamIdentity)).ToArray();
        if (matches.Length != 1) {
            throw new NotSupportedException("The selected image payload could not be resolved uniquely for a safe move.");
        }

        PdfExtractedImage image = matches[0];
        if (!image.IsImageFile) {
            throw new NotSupportedException("Moving this image is not supported because its decoded payload is not a complete portable image file.");
        }
        if (image.IsImageMask) {
            throw new NotSupportedException("Moving an ImageMask placement is not supported because its paint color is owned by page graphics state.");
        }
        if (image.HasUnresolvedTransparencyMask) {
            throw new NotSupportedException("Moving this image is not supported because its PDF transparency mask could not be represented in the extracted image payload.");
        }
        if (string.Equals(image.Filter, "DCTDecode", StringComparison.Ordinal)) {
            if (image.HasExplicitDecode || image.HasDecodeParameters) {
                throw new NotSupportedException("Moving this JPEG image is not supported because PDF Decode or DecodeParms semantics would be lost during restamping.");
            }
            if (image.BitsPerComponent != 8 ||
                (!string.Equals(image.ColorSpace, "DeviceGray", StringComparison.Ordinal) &&
                 !string.Equals(image.ColorSpace, "DeviceRGB", StringComparison.Ordinal) &&
                 !string.Equals(image.ColorSpace, "DeviceCMYK", StringComparison.Ordinal))) {
                throw new NotSupportedException("Moving this JPEG image is not supported because its PDF color-space or component semantics are not portable to image stamping.");
            }
        }
        return image;
    }

    private static ImageTransform ResolvePortableTransform(PdfImagePlacement placement) {
        if (placement.ClipPath != null) {
            throw new NotSupportedException("Replacing or moving a clipped image placement is not supported because the clip path cannot be preserved by image stamping.");
        }
        if (placement.ImageOpacity.HasValue && Math.Abs(placement.ImageOpacity.Value - 1D) > TransformTolerance) {
            throw new NotSupportedException("Replacing or moving an image placement with graphics-state opacity is not supported because the opacity cannot be preserved by image stamping.");
        }
        if (placement.BlendMode.HasValue && placement.BlendMode.Value != OfficeBlendMode.Normal) {
            throw new NotSupportedException("Replacing or moving an image placement with a non-normal blend mode is not supported because the blend state cannot be preserved by image stamping.");
        }
        if (placement.HasSoftMask) {
            throw new NotSupportedException("Replacing or moving an image placement with a graphics-state soft mask is not supported because the mask cannot be preserved by image stamping.");
        }

        double width = Math.Sqrt((placement.A * placement.A) + (placement.B * placement.B));
        double height = Math.Sqrt((placement.C * placement.C) + (placement.D * placement.D));
        double determinant = (placement.A * placement.D) - (placement.B * placement.C);
        double dot = (placement.A * placement.C) + (placement.B * placement.D);
        double scale = Math.Max(1D, width * height);
        if (width <= TransformTolerance || height <= TransformTolerance || determinant <= TransformTolerance || Math.Abs(dot) > TransformTolerance * scale) {
            throw new NotSupportedException("Replacing or moving a skewed, reflected, or degenerate image placement is not supported because its transform cannot be reproduced by image stamping.");
        }

        double rotationDegrees = Math.Atan2(placement.B, placement.A) * (180D / Math.PI);
        return new ImageTransform(width, height, rotationDegrees);
    }

    private static void EnsureRemovablePlacement(PdfImagePlacement placement) {
        if (placement.InlineImageStream != null) {
            throw new NotSupportedException("Editing an inline-image placement is not supported because its binary content-stream framing cannot yet be rewritten independently. XObject image placements remain fully supported.");
        }
        if (placement.Width <= CoordinateTolerance || placement.Height <= CoordinateTolerance) {
            throw new NotSupportedException("Editing a degenerate image placement with zero visible area is not supported.");
        }
    }

    private static void EnsurePlacementRemoved(byte[] pdf, PdfImagePlacement placement, PdfReadOptions? readOptions) {
        if (Placements(pdf, readOptions).Any(candidate => SamePlacementIdentity(candidate, placement))) {
            throw new InvalidOperationException("The selected image placement could not be removed safely; no successful edit result was produced.");
        }
    }

    private static PdfImageStampOptions CreateStampOptions(int pageNumber, double x, double y, double width, double height, double rotationDegrees, PdfImageEditOptions options) =>
        new PdfImageStampOptions {
            PageNumbers = new[] { pageNumber },
            X = x,
            Y = y,
            Width = width,
            Height = height,
            RotationDegrees = rotationDegrees,
            BehindContent = options.Layer == PdfImageEditLayer.BehindExistingContent
        };

    private static bool SamePlacementIdentity(PdfImagePlacement left, PdfImagePlacement right) =>
        left.PageNumber == right.PageNumber &&
        string.Equals(left.ResourceName, right.ResourceName, StringComparison.Ordinal) &&
        left.ObjectNumber == right.ObjectNumber &&
        NearlyEqual(left.A, right.A) && NearlyEqual(left.B, right.B) &&
        NearlyEqual(left.C, right.C) && NearlyEqual(left.D, right.D) &&
        NearlyEqual(left.E, right.E) && NearlyEqual(left.F, right.F) &&
        NearlyEqual(left.X, right.X) && NearlyEqual(left.Y, right.Y) &&
        NearlyEqual(left.Width, right.Width) && NearlyEqual(left.Height, right.Height);

    private static bool Intersects(PdfPageRegion region, PdfImagePlacement placement) =>
        region.PageNumber == placement.PageNumber &&
        region.X < placement.X + placement.Width - CoordinateTolerance &&
        region.Right > placement.X + CoordinateTolerance &&
        region.Y < placement.Y + placement.Height - CoordinateTolerance &&
        region.Top > placement.Y + CoordinateTolerance;

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= CoordinateTolerance;

    private static void ValidatePage(int pageNumber, int pageCount, string parameterName) {
        if (pageNumber > pageCount) throw new ArgumentOutOfRangeException(parameterName, "Page number is outside the current PDF document.");
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(parameterName, "Image movement offset must be finite.");
    }

    internal sealed class ImageMutationResult {
        internal ImageMutationResult(byte[] bytes, int affectedCount) {
            Bytes = bytes;
            AffectedCount = affectedCount;
        }

        internal byte[] Bytes { get; }
        internal int AffectedCount { get; }
    }

    private readonly struct ImageTransform {
        internal ImageTransform(double width, double height, double rotationDegrees) {
            Width = width;
            Height = height;
            RotationDegrees = rotationDegrees;
        }

        internal double Width { get; }
        internal double Height { get; }
        internal double RotationDegrees { get; }
    }
}
