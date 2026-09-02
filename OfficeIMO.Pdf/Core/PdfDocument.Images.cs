using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>One image source used to create a PDF page.</summary>
public sealed class PdfImageDocumentSource {
    private readonly byte[] _bytes;

    /// <summary>Creates an image source from an encoded image payload.</summary>
    public PdfImageDocumentSource(byte[] bytes, string? name = null) {
        Guard.NotNull(bytes, nameof(bytes));
        if (bytes.Length == 0) throw new ArgumentException("Image bytes cannot be empty.", nameof(bytes));
        _bytes = (byte[])bytes!.Clone();
        Name = string.IsNullOrWhiteSpace(name) ? null : name!.Trim();
    }

    /// <summary>Optional source name used as PDF alternate text.</summary>
    public string? Name { get; }

    /// <summary>Returns a caller-owned copy of the encoded image payload.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();

    /// <summary>Creates an image source from a file snapshot.</summary>
    public static PdfImageDocumentSource FromFile(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Image path cannot be empty.", nameof(path));
        string fullPath = Path.GetFullPath(path);
        return new PdfImageDocumentSource(File.ReadAllBytes(fullPath), Path.GetFileName(fullPath));
    }
}

/// <summary>Options for creating one PDF page per source image.</summary>
public sealed class PdfImageDocumentOptions {
    /// <summary>
    /// Optional fixed paper size. When omitted, each page uses the image physical dimensions derived from its DPI.
    /// Images without usable dimensions fall back to <see cref="FallbackPageSize"/>.
    /// </summary>
    public PageSize? FixedPageSize { get; set; }

    /// <summary>Fallback paper size for an image whose physical dimensions cannot be determined.</summary>
    public PageSize FallbackPageSize { get; set; } = PageSizes.A4;

    /// <summary>Automatically rotates a fixed or fallback paper size to match the image orientation.</summary>
    public bool AutoOrientPage { get; set; } = true;

    /// <summary>Uniform page margin in points.</summary>
    public double Margin { get; set; }

    /// <summary>How each image is fitted into the printable page rectangle.</summary>
    public OfficeImageFit Fit { get; set; } = OfficeImageFit.Contain;

    /// <summary>Maximum generated page width or height in points.</summary>
    public double MaximumPageDimension { get; set; } = 14_400D;

    internal PdfImageDocumentOptions CloneAndValidate() {
        if (Margin < 0D || double.IsNaN(Margin) || double.IsInfinity(Margin)) {
            throw new ArgumentOutOfRangeException(nameof(Margin));
        }
#pragma warning disable CA2263 // Generic Enum.IsDefined is unavailable on net472.
        if (!Enum.IsDefined(typeof(OfficeImageFit), Fit)) throw new ArgumentOutOfRangeException(nameof(Fit));
#pragma warning restore CA2263
        if (MaximumPageDimension <= 0D || double.IsNaN(MaximumPageDimension) || double.IsInfinity(MaximumPageDimension)) {
            throw new ArgumentOutOfRangeException(nameof(MaximumPageDimension));
        }
        ValidatePageSize(FallbackPageSize, nameof(FallbackPageSize), validatePrintableArea: false);
        if (FixedPageSize.HasValue) ValidatePageSize(FixedPageSize.Value, nameof(FixedPageSize), validatePrintableArea: true);
        return new PdfImageDocumentOptions {
            FixedPageSize = FixedPageSize,
            FallbackPageSize = FallbackPageSize,
            AutoOrientPage = AutoOrientPage,
            Margin = Margin,
            Fit = Fit,
            MaximumPageDimension = MaximumPageDimension
        };
    }

    private void ValidatePageSize(PageSize size, string name, bool validatePrintableArea) {
        if (size.Width > MaximumPageDimension || size.Height > MaximumPageDimension) {
            throw new ArgumentOutOfRangeException(name, $"Page dimensions cannot exceed {MaximumPageDimension:N0} points.");
        }
        if (validatePrintableArea && (size.Width <= Margin * 2D || size.Height <= Margin * 2D)) {
            throw new ArgumentException("Page margins leave no printable image area.", name);
        }
    }
}

public sealed partial class PdfDocument {
    /// <summary>Creates one PDF page per image file in caller order.</summary>
    public static PdfDocument CreateFromImages(
        IEnumerable<string> imagePaths,
        PdfImageDocumentOptions? options = null) {
        Guard.NotNull(imagePaths, nameof(imagePaths));
        return CreateFromImages(imagePaths.Select(PdfImageDocumentSource.FromFile), options);
    }

    /// <summary>Creates one PDF page per encoded image source in caller order.</summary>
    public static PdfDocument CreateFromImages(
        IEnumerable<PdfImageDocumentSource> images,
        PdfImageDocumentOptions? options = null) {
        return CreateFromImages(images, options, CancellationToken.None);
    }

    /// <summary>Creates one PDF page per encoded image source in caller order with cooperative cancellation.</summary>
    public static PdfDocument CreateFromImages(
        IEnumerable<PdfImageDocumentSource> images,
        PdfImageDocumentOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(images, nameof(images));
        cancellationToken.ThrowIfCancellationRequested();
        var sourceList = new List<PdfImageDocumentSource>();
        foreach (PdfImageDocumentSource source in images) {
            cancellationToken.ThrowIfCancellationRequested();
            sourceList.Add(source);
        }
        PdfImageDocumentSource[] sources = sourceList.ToArray();
        if (sources.Length == 0) throw new ArgumentException("At least one image is required.", nameof(images));
        if (sources.Any(static source => source is null)) {
            throw new ArgumentException("Image sources cannot contain null entries.", nameof(images));
        }

        PdfImageDocumentOptions effective = (options ?? new PdfImageDocumentOptions()).CloneAndValidate();
        var document = new PdfDocument();
        foreach (PdfImageDocumentSource source in sources) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] sourceBytes = source.GetBytes();
            PreparedImage prepared = PrepareImageDocumentSource(sourceBytes);
            cancellationToken.ThrowIfCancellationRequested();
            PageSize pageSize = ResolveImagePageSize(prepared.Info, effective);
            double frameWidth = pageSize.Width - effective.Margin * 2D;
            double frameHeight = pageSize.Height - effective.Margin * 2D;
            string? alternativeText = string.IsNullOrWhiteSpace(source.Name) ? null : source.Name;
            OfficeImageFit placementFit = effective.FixedPageSize.HasValue
                ? effective.Fit
                : OfficeImageFit.Stretch;

            document.AddComposedPage(page => page
                .Size(pageSize)
                .Margin(0D)
                .Canvas(canvas => canvas.Image(
                    prepared.Data,
                    effective.Margin,
                    effective.Margin,
                    frameWidth,
                    frameHeight,
                    new PdfImageStyle { Fit = placementFit },
                    alternativeText: alternativeText)));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return document;
    }

    private static PreparedImage PrepareImageDocumentSource(byte[] sourceBytes) {
        if (!OfficeImageOrientationNormalizer.TryRead(sourceBytes, out OfficeImageOrientation orientation) ||
            orientation == OfficeImageOrientation.Normal) {
            return PrepareImageBytes(sourceBytes);
        }
        if (!OfficeImageOrientationNormalizer.TryNormalizeToPng(
                sourceBytes,
                applyEmbeddedOrientation: true,
                out byte[] normalizedPng,
                out OfficeImageInfo? normalizedInfo) ||
            normalizedInfo == null) {
            throw new NotSupportedException(
                "The image has embedded orientation metadata that could not be normalized safely for PDF composition.");
        }
        return PrepareImageBytes(normalizedPng);
    }

    private static PageSize ResolveImagePageSize(OfficeImageInfo info, PdfImageDocumentOptions options) {
        bool isLandscape = info.Width > info.Height && info.Height > 0;
        PageSize pageSize;
        if (options.FixedPageSize.HasValue) {
            pageSize = options.FixedPageSize.Value;
        } else if (info.Width > 0 && info.Height > 0) {
            double dpiX = ResolveImageDpi(info.DpiX);
            double dpiY = ResolveImageDpi(info.DpiY);
            double width = info.Width * 72D / dpiX;
            double height = info.Height * 72D / dpiY;
            double maximumContentDimension = options.MaximumPageDimension - options.Margin * 2D;
            if (maximumContentDimension <= 0D) {
                throw new ArgumentException("Page margins leave no printable image area.", nameof(options));
            }
            double scale = Math.Min(1D, maximumContentDimension / Math.Max(width, height));
            pageSize = new PageSize(width * scale + options.Margin * 2D, height * scale + options.Margin * 2D);
            return pageSize;
        } else {
            pageSize = options.FallbackPageSize;
        }

        if (options.AutoOrientPage) pageSize = isLandscape ? pageSize.Landscape() : pageSize.Portrait();
        if (pageSize.Width > options.MaximumPageDimension || pageSize.Height > options.MaximumPageDimension) {
            throw new ArgumentOutOfRangeException(nameof(options), "Resolved page dimensions exceed the configured maximum.");
        }
        if (pageSize.Width <= options.Margin * 2D || pageSize.Height <= options.Margin * 2D) {
            throw new ArgumentException("Page margins leave no printable image area.", nameof(options));
        }
        return pageSize;
    }

    private static double ResolveImageDpi(double dpi) =>
        dpi >= 10D && !double.IsNaN(dpi) && !double.IsInfinity(dpi) ? dpi : 96D;
}
