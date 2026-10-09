using System;
using System.Threading;

namespace OfficeIMO.Drawing {
    public static partial class OfficeRasterResampler {
        /// <summary>Captures fit, pixel rounding, crop and peak storage before allocating raster pixels.</summary>
        public static OfficeRasterResizePlan PlanResize(int sourceWidth, int sourceHeight,
            OfficeRasterResizeOptions options, CancellationToken cancellationToken = default) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRasterGuards.EnsureOutputPixels(sourceWidth, sourceHeight, "Source dimensions exceed the managed image limit.");
            int? requestedWidth = options.Width, requestedHeight = options.Height;
            OfficeImageFit fit = options.Fit;
            OfficeRasterResamplingMode mode = options.ResamplingMode;
            OfficeRasterResamplingColorSpace colorSpace = options.ColorSpace;
            if (!requestedWidth.HasValue && !requestedHeight.HasValue) throw new ArgumentException("At least one resize dimension is required.", nameof(options));
            if (requestedWidth <= 0 || requestedHeight <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Resize dimensions must be positive.");
            if (fit == OfficeImageFit.Cover && (!requestedWidth.HasValue || !requestedHeight.HasValue)) throw new ArgumentException("Cover requires both resize dimensions.", nameof(options));
            if (colorSpace != OfficeRasterResamplingColorSpace.EncodedSrgb && colorSpace != OfficeRasterResamplingColorSpace.LinearLight) throw new ArgumentOutOfRangeException(nameof(options), "Unsupported filtering color space.");

            double targetWidth = requestedWidth ?? sourceWidth, targetHeight = requestedHeight ?? sourceHeight;
            if (fit == OfficeImageFit.Contain) {
                if (!requestedWidth.HasValue) targetWidth = sourceWidth * targetHeight / sourceHeight;
                if (!requestedHeight.HasValue) targetHeight = sourceHeight * targetWidth / sourceWidth;
            }
            OfficeImageRenderPlan geometry = OfficeImageRenderPlan.CreateTopLeft(sourceWidth, sourceHeight,
                0, 0, targetWidth, targetHeight, fit);
            int resizeWidth = QuantizeResizeDimension(geometry.ImagePlacement.Width, fit == OfficeImageFit.Cover);
            int resizeHeight = QuantizeResizeDimension(geometry.ImagePlacement.Height, fit == OfficeImageFit.Cover);
            if (fit == OfficeImageFit.Contain) {
                if (requestedWidth.HasValue) resizeWidth = Math.Min(resizeWidth, requestedWidth.Value);
                if (requestedHeight.HasValue) resizeHeight = Math.Min(resizeHeight, requestedHeight.Value);
            }
            int width = fit == OfficeImageFit.Cover ? requestedWidth!.Value : resizeWidth;
            int height = fit == OfficeImageFit.Cover ? requestedHeight!.Value : resizeHeight;
            OfficeRasterGuards.EnsureOutputPixels(width, height, "Resize output dimensions exceed the managed image limit.");
            long workingBytes = GetResizeWorkingSetBytes(sourceWidth, sourceHeight, resizeWidth, resizeHeight, mode, cancellationToken);
            if (resizeWidth != width || resizeHeight != height) {
                long cropPeak = checked(((long)sourceWidth * sourceHeight + (long)resizeWidth * resizeHeight + (long)width * height) * 4L + 72L + 65536L);
                workingBytes = Math.Max(workingBytes, cropPeak);
            }
            if (workingBytes > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Resize intermediates exceed the managed image working-set limit.", nameof(options));
            return new OfficeRasterResizePlan(sourceWidth, sourceHeight, resizeWidth, resizeHeight, width, height, fit, mode, colorSpace, workingBytes);
        }

        /// <summary>Returns independently resized pixels using requested fit bounds and sampling.</summary>
        public static OfficeRasterImage Resize(OfficeRasterImage source, OfficeRasterResizeOptions options,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            return Resize(source, PlanResize(source.Width, source.Height, options, cancellationToken), cancellationToken);
        }

        /// <summary>Executes an immutable resize plan against its required source dimensions, without modifying the source.</summary>
        public static OfficeRasterImage Resize(OfficeRasterImage source, OfficeRasterResizePlan plan,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            cancellationToken.ThrowIfCancellationRequested();
            if (source.Width != plan.SourceWidth || source.Height != plan.SourceHeight) throw new ArgumentException("Source dimensions do not match the resize plan.", nameof(source));
            OfficeRasterImage resized = Resize(source, plan.ResizeWidth, plan.ResizeHeight, plan.ResamplingMode, plan.ColorSpace, cancellationToken);
            return plan.ResizeWidth == plan.Width && plan.ResizeHeight == plan.Height ? resized :
                OfficeRasterTransforms.Crop(resized, plan.CropX, plan.CropY, plan.Width, plan.Height, cancellationToken);
        }

        private static int QuantizeResizeDimension(double value, bool roundUp) {
            double rounded = roundUp ? Math.Ceiling(value) : Math.Round(value, MidpointRounding.AwayFromZero);
            if (double.IsNaN(rounded) || double.IsInfinity(rounded) || rounded > int.MaxValue) throw new ArgumentException("Resize dimensions are not representable.");
            return Math.Max(1, (int)rounded);
        }
    }
}
