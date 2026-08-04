using OfficeIMO.AsciiDoc;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Latex;
using OfficeIMO.Markdown;
using OfficeIMO.OneNote;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Rtf;
using OfficeIMO.Visio;
using OfficeIMO.Word;

internal static class ConversionApiCompileContract {
    // This method is intentionally never executed. Its body makes every representative
    // catalog entry a compile-time dependency, so generated API guidance cannot drift
    // into a plausible-looking but invalid signature.
    internal static void Verify(
        Stream stream,
        string source,
        AsciiDocDocument asciiDoc,
        ExcelDocument excel,
        HtmlConversionDocument html,
        LatexDocument latex,
        MarkdownDoc markdown,
        MhtmlDocument mhtml,
        OneNoteSection oneNote,
        OdtDocument odt,
        OdsDocument ods,
        OdpPresentation odp,
        PdfDocument pdf,
        PowerPointPresentation powerPoint,
        RtfDocument rtf,
        RtfReadResult rtfRead,
        VisioDocument visio,
        WordDocument word) {
        _ = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(source);
        _ = OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownConverterExtensions.ToMarkdownDocumentResult(asciiDoc);
        _ = OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownConverterExtensions.ToAsciiDocDocumentResult(markdown);
        _ = OfficeIMO.AsciiDoc.Pdf.AsciiDocPdfConverterExtensions.ToPdfDocumentResult(asciiDoc);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToMarkdownDocumentResult(latex);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToLatexDocumentResult(markdown);
        _ = OfficeIMO.Latex.Pdf.LatexPdfConverterExtensions.ToPdfDocumentResult(latex);
        _ = OfficeIMO.Markdown.Html.HtmlMarkdownConverterExtensions.ToMarkdownDocumentResult(html);
        _ = OfficeIMO.Markdown.Pdf.MarkdownPdfConverterExtensions.ToPdfDocumentResult(markdown);
        _ = OfficeIMO.Rtf.Markdown.RtfMarkdownConverterExtensions.ToMarkdownResult(rtf);
        _ = OfficeIMO.Rtf.Markdown.RtfMarkdownConverterExtensions.ToRtfDocumentResult(markdown);
        _ = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions.ToMarkdownDocumentResult(word);
        _ = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions.ToWordDocumentResult(markdown);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionExtensions.ToOpenDocumentResult(word);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionExtensions.ToWordDocumentResult(odt);
        _ = OfficeIMO.Html.HtmlRtfConverterExtensions.ToRtfDocumentResult(html);
        _ = OfficeIMO.Html.HtmlRtfConverterExtensions.ToHtmlResult(rtf);
        _ = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToRtfDocumentResult(word);
        _ = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToWordDocumentResult(rtfRead);
        _ = OfficeIMO.Excel.Html.ExcelHtmlConverterExtensions.ToHtmlResult(excel);
        _ = OfficeIMO.Excel.Html.HtmlExcelConverterExtensions.ToExcelDocumentResult(html);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionExtensions.ToOpenDocumentResult(excel);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionExtensions.ToExcelDocumentResult(ods);
        _ = OfficeIMO.Word.Html.WordHtmlConverterExtensions.ToHtmlResult(word);
        _ = OfficeIMO.Word.Html.WordHtmlConverterExtensions.ToWordDocumentResult(html);
        _ = OfficeIMO.PowerPoint.Html.PowerPointHtmlConverterExtensions.ToHtmlResult(powerPoint);
        _ = OfficeIMO.PowerPoint.Html.HtmlPowerPointConverterExtensions.ToPowerPointPresentationResult(html);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionExtensions.ToOpenDocumentResult(powerPoint);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionExtensions.ToPowerPointPresentationResult(odp);
        _ = OfficeIMO.OneNote.Markdown.OneNoteMarkdownConverterExtensions.ToMarkdownDocumentResult(oneNote);
        _ = OfficeIMO.OneNote.Html.OneNoteHtmlConverterExtensions.ToHtmlDocumentResult(oneNote);
        _ = OfficeIMO.OneNote.Html.HtmlOneNoteConverterExtensions.ToOneNoteSectionResult(html);
        _ = OfficeIMO.OneNote.Pdf.OneNoteSectionPdfConverterExtensions.ToPdfDocumentResult(oneNote);
        _ = OfficeIMO.Word.Pdf.WordPdfConverterExtensions.ToPdfDocumentResult(word);
        _ = OfficeIMO.Word.Pdf.PdfWordConverterExtensions.ToWordDocumentResult(pdf);
        _ = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdfDocumentResult(excel);
        _ = OfficeIMO.Excel.Pdf.PdfExcelTableConverterExtensions.ImportTablesToExcelDocumentResult(pdf);
        _ = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPdfDocumentResult(powerPoint);
        _ = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPowerPointPresentationResult(pdf);
        _ = OfficeIMO.Html.Pdf.HtmlPdfConverterExtensions.ToPdfDocumentResult(html);
        _ = OfficeIMO.Html.Pdf.HtmlPdfConverterExtensions.ToPdfDocumentResult(mhtml);
        _ = OfficeIMO.Html.Pdf.PdfHtmlConverterExtensions.ToHtmlResult(pdf);
        _ = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToPdfDocumentResult(rtf);
        _ = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToRtfDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfConversionExtensions.ToPdfDocumentResult(odt);
        _ = OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfConversionExtensions.ToOdtDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfConversionExtensions.ToPdfDocumentResult(ods);
        _ = OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfConversionExtensions.ToOdsDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfConversionExtensions.ToPdfDocumentResult(odp);
        _ = OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfConversionExtensions.ToOdpPresentationResult(pdf);
        _ = OfficeIMO.Visio.Pdf.VisioPdfConverterExtensions.ToPdfDocumentResult(visio);
    }
}
