namespace OfficeIMO.Pdf;

/// <summary>Controls arbitrary visual canvas content stamped onto existing PDF pages.</summary>
public sealed class PdfCanvasStampOptions {
    private double _opacity = 1D;

    /// <summary>Optional target-page selector. Null applies the canvas callback to every page.</summary>
    public PdfPageSelector? TargetPages { get; set; }

    /// <summary>Places generated visual content before the existing page content streams.</summary>
    public bool BehindContent { get; set; }

    /// <summary>Canvas opacity from zero through one.</summary>
    public double Opacity {
        get => _opacity;
        set {
            if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(Opacity), value, "Canvas stamp opacity must be finite and between zero and one.");
            }

            _opacity = value;
        }
    }

    /// <summary>Sets target pages from a rich page-selector expression.</summary>
    public PdfCanvasStampOptions UseTargetPages(string selector) {
        TargetPages = PdfPageSelector.Parse(selector);
        return this;
    }
}
