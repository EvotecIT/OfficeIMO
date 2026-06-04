using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable image placement and rhythm style.
/// </summary>
public sealed class PdfImageStyle {
    private PdfAlign _align = PdfAlign.Left;
    private OfficeImageFit _fit = OfficeImageFit.Stretch;
    private OfficeClipPath? _clipPath;
    private PdfImageSourceCrop? _sourceCrop;
    private double _spacingBefore;
    private double _spacingAfter;
    private string? _alternativeText;

    /// <summary>Image alignment within the current content frame.</summary>
    public PdfAlign Align {
        get => _align;
        set {
            Guard.LeftCenterRightAlign(value, nameof(Align), "Image");
            _align = value;
        }
    }

    /// <summary>How the image is fitted inside its target box.</summary>
    public OfficeImageFit Fit {
        get => _fit;
        set {
            PdfDocument.ValidateImageFit(value, nameof(Fit));
            _fit = value;
        }
    }

    /// <summary>Optional clipping path applied inside the image target box.</summary>
    public OfficeClipPath? ClipPath {
        get => _clipPath?.Clone();
        set => _clipPath = value?.Clone();
    }

    /// <summary>Optional source crop applied before fitting the image into the target box.</summary>
    public PdfImageSourceCrop? SourceCrop {
        get => _sourceCrop?.Clone();
        set => _sourceCrop = value?.Clone();
    }

    /// <summary>Vertical space before the image in the surrounding document flow, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Image spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the image in the surrounding document flow, in points.</summary>
    public double SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingAfter), "Image spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Moves the image to the next page with the first visible part of the following block when they fit together.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>When true, oversized flow images are proportionally reduced to fit the current page or column frame.</summary>
    public bool ScaleDownToFit { get; set; }

    /// <summary>Optional alternate text for meaningful generated images.</summary>
    public string? AlternativeText {
        get => _alternativeText;
        set {
            if (value != null) {
                Guard.NotNullOrWhiteSpace(value, nameof(AlternativeText));
            }

            _alternativeText = value;
        }
    }

    /// <summary>Creates a copy of this image style.</summary>
    public PdfImageStyle Clone() {
        return new PdfImageStyle {
            Align = Align,
            Fit = Fit,
            ClipPath = _clipPath,
            SourceCrop = _sourceCrop,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            KeepWithNext = KeepWithNext,
            ScaleDownToFit = ScaleDownToFit,
            AlternativeText = AlternativeText
        };
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }
}
