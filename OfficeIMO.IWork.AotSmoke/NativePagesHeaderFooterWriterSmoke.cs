using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class NativePagesHeaderFooterWriterSmoke {
    internal static void Run() {
        var document = IWorkPagesDocument.Create();
        document.SetHeader(new IWorkTextRunData("OfficeIMO — Pages header", "Arial", 14, "003366", bold: true), 30);
        document.SetFooter(new IWorkTextRunData("Zażółć café — footer", "Arial", 12, "555555", italic: true), 36);
        document.AddParagraph("Native header and footer creation", fontSizePoints: 20);
        for (int index = 1; index <= 72; index++)
            document.AddParagraph("Body line " + index.ToString("D2", CultureInfo.InvariantCulture) + " — header/footer creation");
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == "9A8A0DC6ACE759E222054A2B7365A53DDA88CC561C8706322645E5B92C2B66F2",
            "Pages header/footer creation differs from the independent Apple-accepted graph.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "Page-text creation is not deterministic.");
        Verify(bytes, "OfficeIMO — Pages header", "Zażółć café — footer");
        string source = Path.Combine(AppContext.BaseDirectory, "pages-created-headers-edited.pages");
        byte[] native = File.ReadAllBytes(source);
        Require(Convert.ToHexString(SHA256.HashData(native)) == "646B208C135743B886353BC8A63264EBE4B1F4981300D9C88FB7CDF018C8C16F",
            "The native-edited page-text reference changed.");
        Verify(native, "OfficeIMO — Native header", "Zażółć café — native footer");
        using var output = new MemoryStream(new byte[] { 9, 8, 7 });
        try { document.Save(output, new() { MaximumParagraphs = 74 }); throw new InvalidOperationException("Header/footer paragraphs escaped their budget."); }
        catch (InvalidDataException) { Require(output.ToArray().SequenceEqual(new byte[] { 9, 8, 7 }), "Budget rejection mutated output."); }
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        try { document.Save(output, cancellationToken: cancelled.Token); throw new InvalidOperationException("Cancelled page-text save succeeded."); }
        catch (OperationCanceledException) { Require(output.ToArray().SequenceEqual(new byte[] { 9, 8, 7 }), "Cancelled save mutated output."); }
        foreach (Action setEmpty in new Action[] {
            () => document.SetHeader(new IWorkTextRunData("", "Times New Roman", 24, "003366", bold: true)),
            () => document.SetFooter("") }) {
            try { setEmpty(); throw new InvalidOperationException("Zero-length page text was accepted."); }
            catch (ArgumentException) { Require(bytes.SequenceEqual(document.SaveBytes()), "Rejected page text mutated the model."); }
        }
        document.ClearHeader(); document.ClearFooter();
        Require(document.Header == null && document.Footer == null && document.HeaderDistancePoints == 0 && document.FooterDistancePoints == 0,
            "Clearing page text retained enabled state or distances.");
        Console.WriteLine("PASS | native Pages header/footer creation | deterministic default left text, explicit distances/styles, Apple edits, strict saved/reopened DOCX, shared paragraph limits and cancellation");
    }

    private static void Verify(byte[] bytes, string header, string footer) {
        IWorkPagesProjection projection = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages();
        Require(projection.Diagnostics.Count == 0 && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0,
            "Page text has incomplete declarations or references.");
        Require(projection.Sections.Count == 1 && projection.Sections[0].DefaultPageHeaderContents.Single().PlainText == header
            && projection.Sections[0].DefaultPageFooterContents.Single().PlainText == footer
            && projection.PageLayout!.HeaderMarginPoints == 30 && projection.PageLayout.FooterMarginPoints == 36,
            "Page text or distances differ.");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes), conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Require(!new OpenXmlValidator().Validate(saved).Any(), "Saved page-text DOCX is invalid.");
        W.Paragraph h = saved.MainDocumentPart!.HeaderParts.Single().Header!.Elements<W.Paragraph>().Single();
        W.Paragraph f = saved.MainDocumentPart.FooterParts.Single().Footer!.Elements<W.Paragraph>().Single();
        Require(h.InnerText == header && f.InnerText == footer, "Saved page text changed.");
        W.Run hr = h.Elements<W.Run>().Single(run => run.InnerText.Length > 0), fr = f.Elements<W.Run>().Single(run => run.InnerText.Length > 0);
        Require(hr.RunProperties!.Bold != null && (hr.RunProperties.Bold.Val?.Value ?? true)
            && fr.RunProperties!.Italic != null && (fr.RunProperties.Italic.Val?.Value ?? true), "Saved page-text decorations changed.");
    }

    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
