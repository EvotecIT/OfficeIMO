using OfficeIMO.Chm;
using OfficeIMO.DjVu.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Visio.Pdf;
using OfficeIMO.Xps;

namespace OfficeIMO.Workflows;

/// <summary>Settings for the additional format routes registered by <see cref="OfficeWorkflow.RegisterConversions"/>.</summary>
public sealed class OfficeWorkflowRouteSettings : IOfficeConversionSettings {
    /// <summary>RTF projection and PDF settings.</summary>
    public RtfToPdfOptions? Rtf { get; set; }
    /// <summary>OpenDocument drawing projection and PDF settings.</summary>
    public OdgToPdfOptions? Draw { get; set; }
    /// <summary>Visio drawing projection and PDF settings.</summary>
    public VisioToPdfOptions? Visio { get; set; }
    /// <summary>Publisher decoding limits.</summary>
    public PublisherReadOptions? PublisherRead { get; set; }
    /// <summary>Publisher PDF output settings.</summary>
    public PdfCreateOptions? PublisherPdf { get; set; }
    /// <summary>XPS reading and PDF settings.</summary>
    public XpsToPdfOptions? Xps { get; set; }
    /// <summary>DjVu page and PDF settings.</summary>
    public DjVuToPdfOptions? DjVu { get; set; }
    /// <summary>CHM topic selection and aggregate conversion limits.</summary>
    public ChmConversionOptions? Chm { get; set; }
    /// <summary>HTML rendering settings for CHM topics.</summary>
    public HtmlToPdfOptions? ChmHtml { get; set; }

    internal OfficeWorkflowRenderingOptions Capture(bool failOnLoss) => new OfficeWorkflowRenderingOptions {
        Rtf = Rtf, Draw = Draw, Visio = Visio, PublisherRead = PublisherRead, PublisherPdf = PublisherPdf,
        Xps = Xps, DjVu = DjVu, Chm = Chm, Html = ChmHtml, RequireNoLoss = failOnLoss
    }.Clone();

    /// <inheritdoc />
    public IOfficeConversionSettings Snapshot() {
        OfficeWorkflowRenderingOptions copy = Capture(false);
        return new OfficeWorkflowRouteSettings { Rtf = copy.Rtf, Draw = copy.Draw, Visio = copy.Visio,
            PublisherRead = copy.PublisherRead, PublisherPdf = copy.PublisherPdf, Xps = copy.Xps,
            DjVu = copy.DjVu, Chm = copy.Chm, ChmHtml = copy.Html };
    }
}
