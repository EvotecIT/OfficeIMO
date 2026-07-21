namespace OfficeIMO.Pdf;

/// <summary>Controls arbitrary visual canvas content stamped onto existing PDF pages.</summary>
public sealed class PdfCanvasStampOptions {
    private double _opacity = 1D;
    private PdfOptions? _renderingOptions;

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

    /// <summary>
    /// Generated-PDF rendering options used for the temporary visual overlay.
    /// The value is cloned on assignment and retrieval. Page geometry, margins, encryption, page chrome, and page decorations are controlled by the stamping operation.
    /// Use this to carry registered embedded fonts, text shaping, image handling, and other visual rendering configuration into canvas content.
    /// Tagged structure mode must remain disabled because a visual page import cannot merge the temporary document's structure tree into the target.
    /// </summary>
    public PdfOptions? RenderingOptions {
        get => _renderingOptions?.Clone();
        set => _renderingOptions = value?.Clone();
    }

    internal PdfOptions? RenderingOptionsSnapshot => _renderingOptions?.Clone();

    /// <summary>Sets target pages from a rich page-selector expression.</summary>
    public PdfCanvasStampOptions UseTargetPages(string selector) {
        TargetPages = PdfPageSelector.Parse(selector);
        return this;
    }

    /// <summary>Uses cloned generated-PDF rendering options for visual canvas content.</summary>
    public PdfCanvasStampOptions UseRenderingOptions(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        RenderingOptions = options;
        return this;
    }
}
