namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRouteSettings {
    /// <summary>Markdown outline parsing, slide boundaries, front matter, notes and native presentation resource policy.</summary>
    public MarkdownToPowerPointOptions? MarkdownPowerPoint { get; set; }
    /// <summary>Markdown parsing, encoding and syntax settings.</summary>
    public OfficeIMO.Markdown.MarkdownLoadOptions? MarkdownLoad { get; set; }

    /// <summary>AsciiDoc parsing limits and native profile.</summary>
    public OfficeIMO.AsciiDoc.AsciiDocParseOptions? AsciiDocLoad { get; set; }

    /// <summary>AsciiDoc to Markdown mapping and loss policy.</summary>
    public OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownExportOptions? AsciiDocMarkdown { get; set; }

    /// <summary>Markdown to AsciiDoc generation and loss policy.</summary>
    public OfficeIMO.AsciiDoc.Markdown.MarkdownAsciiDocExportOptions? MarkdownAsciiDoc { get; set; }

    /// <summary>LaTeX parsing and bounded macro settings.</summary>
    public OfficeIMO.Latex.LatexParseOptions? LatexLoad { get; set; }

    /// <summary>LaTeX to Markdown mapping and loss policy.</summary>
    public OfficeIMO.Latex.Markdown.LatexMarkdownExportOptions? LatexMarkdown { get; set; }

    /// <summary>Markdown to LaTeX generation and loss policy.</summary>
    public OfficeIMO.Latex.Markdown.MarkdownLatexExportOptions? MarkdownLatex { get; set; }

    /// <summary>RTF native parsing and security limits.</summary>
    public OfficeIMO.Rtf.RtfLoadOptions? RtfLoad { get; set; }

    /// <summary>Native RTF output settings.</summary>
    public OfficeIMO.Rtf.RtfSaveOptions? RtfSave { get; set; }

    /// <summary>RTF to HTML conversion policy.</summary>
    public OfficeIMO.Html.RtfHtmlExportOptions? RtfHtml { get; set; }

    /// <summary>HTML to RTF conversion and resource policy.</summary>
    public OfficeIMO.Html.HtmlRtfExportOptions? HtmlRtf { get; set; }

    /// <summary>HTML-to-Word import settings for the HTML→Word→RTF route, including font, base directory and configured stylesheets.</summary>
    /// <remarks>Use instead of HtmlRtf to select the Word import path; WordRtf and RtfSave control its second stage.</remarks>
    public OfficeIMO.Word.Html.HtmlWordExportOptions? HtmlRtfImport { get; set; }

    /// <summary>RTF to Markdown conversion policy.</summary>
    public OfficeIMO.Rtf.Markdown.RtfMarkdownExportOptions? RtfMarkdown { get; set; }

    /// <summary>Markdown to RTF conversion policy.</summary>
    public OfficeIMO.Rtf.Markdown.MarkdownRtfExportOptions? MarkdownRtf { get; set; }

    /// <summary>Word to RTF conversion policy.</summary>
    public OfficeIMO.Word.Rtf.WordRtfExportOptions? WordRtf { get; set; }

    /// <summary>RTF to Word conversion policy.</summary>
    public OfficeIMO.Word.Rtf.RtfWordExportOptions? RtfWord { get; set; }

    /// <summary>PDF to RTF reconstruction policy.</summary>
    public OfficeIMO.Rtf.Pdf.PdfRtfExportOptions? PdfRtf { get; set; }

    /// <summary>ODF package, XML, encryption and resource limits.</summary>
    public OfficeIMO.OpenDocument.OdfLoadOptions? OpenDocumentLoad { get; set; }

    /// <summary>Native ODF serialization, signatures and encryption.</summary>
    public OfficeIMO.OpenDocument.OdfSaveOptions? OpenDocumentSave { get; set; }

    /// <summary>Word and ODT conversion options, including the three loss thresholds.</summary>
    public OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionOptions? WordOpenDocument { get; set; }

    /// <summary>Excel and ODS conversion options, including the three loss thresholds.</summary>
    public OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionOptions? ExcelOpenDocument { get; set; }

    /// <summary>PowerPoint and ODP conversion options, including the three loss thresholds.</summary>
    public OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionOptions? PowerPointOpenDocument { get; set; }

    /// <summary>OneNote native reading and asset limits.</summary>
    public OfficeIMO.OneNote.OneNoteLoadOptions? OneNoteLoad { get; set; }

    /// <summary>OneNote notebook hierarchy, package budgets and section-error policy.</summary>
    public OfficeIMO.OneNote.OneNoteNotebookReaderOptions? OneNoteNotebookLoad { get; set; }

    /// <summary>HTML layout applied to the semantic OneNote projection.</summary>
    public OfficeIMO.Markdown.MarkdownHtmlExportOptions? OneNoteHtmlRendering { get; set; }

    /// <summary>OneNote native section output settings.</summary>
    public OfficeIMO.OneNote.OneNoteSaveOptions? OneNoteSave { get; set; }

    /// <summary>OneNote semantic HTML projection settings.</summary>
    public OfficeIMO.OneNote.Html.OneNoteHtmlExportOptions? OneNoteHtml { get; set; }

    /// <summary>HTML import into a native OneNote section.</summary>
    public OfficeIMO.OneNote.Html.HtmlOneNoteExportOptions? HtmlOneNote { get; set; }

    /// <summary>OneNote semantic Markdown projection settings.</summary>
    public OfficeIMO.OneNote.Markdown.OneNoteMarkdownExportOptions? OneNoteMarkdown { get; set; }

    /// <summary>OneNote semantic PDF projection settings.</summary>
    public OfficeIMO.OneNote.Pdf.OneNotePdfExportOptions? OneNotePdf { get; set; }

    /// <summary>OneNote visual HTML rendering; use instead of semantic HTML settings.</summary>
    public OfficeIMO.OneNote.Html.OneNoteVisualHtmlExportOptions? OneNoteVisualHtml { get; set; }

    /// <summary>OneNote visual PDF rendering; use instead of semantic PDF settings.</summary>
    public OfficeIMO.OneNote.Pdf.OneNoteVisualPdfExportOptions? OneNoteVisualPdf { get; set; }

    /// <summary>AsciiDoc PDF projection settings.</summary>
    public OfficeIMO.AsciiDoc.Pdf.AsciiDocPdfExportOptions? AsciiDocPdf { get; set; }

    /// <summary>LaTeX PDF projection settings.</summary>
    public OfficeIMO.Latex.Pdf.LatexPdfExportOptions? LatexPdf { get; set; }

    /// <summary>ODT PDF projection settings.</summary>
    public OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfExportOptions? OdtPdf { get; set; }

    /// <summary>ODS PDF projection settings.</summary>
    public OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfExportOptions? OdsPdf { get; set; }

    /// <summary>ODP PDF projection settings.</summary>
    public OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfExportOptions? OdpPdf { get; set; }

    /// <summary>Returns an independent snapshot containing only settings accepted by the selected canonical route.</summary>
    public OfficeWorkflowRouteSettings ForRoute(string routeId) {
        var copy = (OfficeWorkflowRouteSettings)Snapshot();
        return routeId switch {
            "markdown-pptx" => new OfficeWorkflowRouteSettings { MarkdownPowerPoint = copy.MarkdownPowerPoint },
            "rtf-pdf" => new OfficeWorkflowRouteSettings { Rtf = copy.Rtf },
            "odg-pdf" => new OfficeWorkflowRouteSettings { Draw = copy.Draw },
            "visio-pdf" => new OfficeWorkflowRouteSettings { Visio = copy.Visio },
            "xps-pdf" => new OfficeWorkflowRouteSettings { Xps = copy.Xps },
            "djvu-pdf" => new OfficeWorkflowRouteSettings { DjVu = copy.DjVu },
            "chm-pdf" => new OfficeWorkflowRouteSettings { Chm = copy.Chm, ChmHtml = copy.ChmHtml },
            "chm-markdown" => new OfficeWorkflowRouteSettings { Chm = copy.Chm },
            "chm-epub" => new OfficeWorkflowRouteSettings { Chm = copy.Chm },
            "publisher-pdf" => new OfficeWorkflowRouteSettings { PublisherRead = copy.PublisherRead, PublisherPdf = copy.PublisherPdf },
            "asciidoc-markdown" => new OfficeWorkflowRouteSettings { AsciiDocLoad = copy.AsciiDocLoad, AsciiDocMarkdown = copy.AsciiDocMarkdown },
            "markdown-asciidoc" => new OfficeWorkflowRouteSettings { MarkdownLoad = copy.MarkdownLoad, MarkdownAsciiDoc = copy.MarkdownAsciiDoc },
            "latex-markdown" => new OfficeWorkflowRouteSettings { LatexLoad = copy.LatexLoad, LatexMarkdown = copy.LatexMarkdown },
            "markdown-latex" => new OfficeWorkflowRouteSettings { MarkdownLoad = copy.MarkdownLoad, MarkdownLatex = copy.MarkdownLatex },
            "rtf-html" => new OfficeWorkflowRouteSettings { RtfLoad = copy.RtfLoad, RtfHtml = copy.RtfHtml },
            "html-rtf" => new OfficeWorkflowRouteSettings { HtmlRtf = copy.HtmlRtf, HtmlRtfImport = copy.HtmlRtfImport, WordRtf = copy.WordRtf, RtfSave = copy.RtfSave },
            "rtf-markdown" => new OfficeWorkflowRouteSettings { RtfLoad = copy.RtfLoad, RtfMarkdown = copy.RtfMarkdown },
            "markdown-rtf" => new OfficeWorkflowRouteSettings { MarkdownLoad = copy.MarkdownLoad, MarkdownRtf = copy.MarkdownRtf, RtfSave = copy.RtfSave },
            "rtf-docx" => new OfficeWorkflowRouteSettings { RtfLoad = copy.RtfLoad, RtfWord = copy.RtfWord },
            "docx-rtf" => new OfficeWorkflowRouteSettings { WordRtf = copy.WordRtf, RtfSave = copy.RtfSave },
            "pdf-rtf" => new OfficeWorkflowRouteSettings { PdfRtf = copy.PdfRtf, RtfSave = copy.RtfSave },
            "docx-odt" => new OfficeWorkflowRouteSettings { WordOpenDocument = copy.WordOpenDocument, OpenDocumentSave = copy.OpenDocumentSave },
            "odt-docx" => new OfficeWorkflowRouteSettings { OpenDocumentLoad = copy.OpenDocumentLoad, WordOpenDocument = copy.WordOpenDocument },
            "xlsx-ods" => new OfficeWorkflowRouteSettings { ExcelOpenDocument = copy.ExcelOpenDocument, OpenDocumentSave = copy.OpenDocumentSave },
            "ods-xlsx" => new OfficeWorkflowRouteSettings { OpenDocumentLoad = copy.OpenDocumentLoad, ExcelOpenDocument = copy.ExcelOpenDocument },
            "pptx-odp" => new OfficeWorkflowRouteSettings { PowerPointOpenDocument = copy.PowerPointOpenDocument, OpenDocumentSave = copy.OpenDocumentSave },
            "odp-pptx" => new OfficeWorkflowRouteSettings { OpenDocumentLoad = copy.OpenDocumentLoad, PowerPointOpenDocument = copy.PowerPointOpenDocument },
            "onenote-html" => new OfficeWorkflowRouteSettings { OneNoteLoad = copy.OneNoteLoad, OneNoteNotebookLoad = copy.OneNoteNotebookLoad, OneNoteMarkdown = copy.OneNoteMarkdown, OneNoteHtml = copy.OneNoteHtml, OneNoteHtmlRendering = copy.OneNoteHtmlRendering, OneNoteVisualHtml = copy.OneNoteVisualHtml },
            "html-onenote" => new OfficeWorkflowRouteSettings { HtmlOneNote = copy.HtmlOneNote, OneNoteSave = copy.OneNoteSave },
            "onenote-markdown" => new OfficeWorkflowRouteSettings { OneNoteLoad = copy.OneNoteLoad, OneNoteNotebookLoad = copy.OneNoteNotebookLoad, OneNoteMarkdown = copy.OneNoteMarkdown },
            "onenote-pdf" => new OfficeWorkflowRouteSettings { OneNoteLoad = copy.OneNoteLoad, OneNoteNotebookLoad = copy.OneNoteNotebookLoad, OneNoteMarkdown = copy.OneNoteMarkdown, OneNotePdf = copy.OneNotePdf, OneNoteVisualPdf = copy.OneNoteVisualPdf },
            "asciidoc-pdf" => new OfficeWorkflowRouteSettings { AsciiDocLoad = copy.AsciiDocLoad, AsciiDocPdf = copy.AsciiDocPdf },
            "latex-pdf" => new OfficeWorkflowRouteSettings { LatexLoad = copy.LatexLoad, LatexPdf = copy.LatexPdf },
            "odt-pdf" => new OfficeWorkflowRouteSettings { OpenDocumentLoad = copy.OpenDocumentLoad, OdtPdf = copy.OdtPdf },
            "ods-pdf" => new OfficeWorkflowRouteSettings { OpenDocumentLoad = copy.OpenDocumentLoad, OdsPdf = copy.OdsPdf },
            "odp-pdf" => new OfficeWorkflowRouteSettings { OpenDocumentLoad = copy.OpenDocumentLoad, OdpPdf = copy.OdpPdf },
            "heic-pdf" or "book-project-epub" => new OfficeWorkflowRouteSettings(),
            _ => throw new ArgumentException("Choose a Workflows conversion route.", nameof(routeId))
        };
    }

    private OfficeWorkflowRouteSettings CopyPairSettingsTo(OfficeWorkflowRouteSettings copy) {
        copy.MarkdownPowerPoint = MarkdownPowerPoint?.Clone();
        copy.MarkdownLoad = MarkdownLoad?.Clone();
        copy.AsciiDocLoad = Snapshot(AsciiDocLoad);
        copy.AsciiDocMarkdown = Snapshot(AsciiDocMarkdown);
        copy.MarkdownAsciiDoc = Snapshot(MarkdownAsciiDoc);
        copy.LatexLoad = Snapshot(LatexLoad);
        copy.LatexMarkdown = Snapshot(LatexMarkdown);
        copy.MarkdownLatex = Snapshot(MarkdownLatex);
        copy.RtfLoad = Snapshot(RtfLoad);
        copy.RtfSave = Snapshot(RtfSave);
        copy.RtfHtml = RtfHtml?.Clone();
        copy.HtmlRtf = HtmlRtf?.Clone();
        copy.HtmlRtfImport = HtmlRtfImport?.Clone();
        copy.RtfMarkdown = RtfMarkdown?.Clone();
        copy.MarkdownRtf = MarkdownRtf?.Clone();
        copy.WordRtf = WordRtf?.Clone();
        copy.RtfWord = RtfWord?.Clone();
        copy.PdfRtf = PdfRtf?.Clone();
        copy.OpenDocumentLoad = Snapshot(OpenDocumentLoad);
        copy.OpenDocumentSave = Snapshot(OpenDocumentSave);
        copy.WordOpenDocument = WordOpenDocument?.Clone();
        copy.ExcelOpenDocument = ExcelOpenDocument?.Clone();
        copy.PowerPointOpenDocument = PowerPointOpenDocument?.Clone();
        copy.OneNoteLoad = Snapshot(OneNoteLoad);
        copy.OneNoteNotebookLoad = Snapshot(OneNoteNotebookLoad);
        copy.OneNoteHtmlRendering = OneNoteHtmlRendering == null ? null : new OfficeIMO.OneNote.Html.OneNoteHtmlExportOptions { HtmlOptions = OneNoteHtmlRendering }.Clone().HtmlOptions;
        copy.OneNoteSave = Snapshot(OneNoteSave);
        copy.OneNoteHtml = OneNoteHtml?.Clone();
        copy.HtmlOneNote = HtmlOneNote?.Clone();
        copy.OneNoteMarkdown = OneNoteMarkdown?.Clone();
        copy.OneNotePdf = OneNotePdf?.Clone();
        copy.OneNoteVisualHtml = OneNoteVisualHtml?.Clone();
        copy.OneNoteVisualPdf = OneNoteVisualPdf?.Clone();
        copy.AsciiDocPdf = AsciiDocPdf?.Clone();
        copy.LatexPdf = LatexPdf?.Clone();
        copy.OdtPdf = OdtPdf?.Clone();
        copy.OdsPdf = OdsPdf?.Clone();
        copy.OdpPdf = OdpPdf?.Clone();
        if (copy.WordRtf != null) copy.WordRtf.SaveOptions = Snapshot(WordRtf?.SaveOptions);
        if (copy.HtmlRtf != null) copy.HtmlRtf.SaveOptions = Snapshot(HtmlRtf?.SaveOptions);
        if (copy.MarkdownRtf != null) copy.MarkdownRtf.SaveOptions = Snapshot(MarkdownRtf?.SaveOptions);
        if (copy.RtfMarkdown != null) {
            copy.RtfMarkdown.MarkdownWriteOptions = RtfMarkdown?.MarkdownWriteOptions?.Clone();
            copy.RtfMarkdown.InlineReaderOptions = RtfMarkdown?.InlineReaderOptions?.Clone();
        }
        if (copy.PdfRtf != null) copy.PdfRtf.ReadOptions = PdfRtf?.ReadOptions?.Clone();
        if (copy.OdtPdf != null) copy.OdtPdf.ConversionOptions = OdtPdf?.ConversionOptions?.Clone();
        if (copy.OdsPdf != null) copy.OdsPdf.ConversionOptions = OdsPdf?.ConversionOptions?.Clone();
        if (copy.OdpPdf != null) copy.OdpPdf.ConversionOptions = OdpPdf?.ConversionOptions?.Clone();
        return copy;
    }
}
