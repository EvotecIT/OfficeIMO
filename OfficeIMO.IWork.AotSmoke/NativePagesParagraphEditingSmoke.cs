using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesParagraphEditingSmoke {
    internal static void Run() {
        var source = Open("pages-paragraph-original.pages"); var body = source.ReadPages().Body;
        var edit = source.BeginEdit(); var id = body.SourceIdentity!;
        edit.SetPagesBodyParagraphText(id, 1, "Edited paragraph layout");
        edit.SetPagesBodyParagraphText(id, 3, "Start\tLeft\tCenter\tRight\t234.56");
        edit.SetPagesBodyParagraphText(id, 7, "Updated reset Native layout edit.");
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        var output = edit.SaveBytesResult(options);
        Require(Convert.ToHexString(SHA256.HashData(output.GetPackageBytes())) == "87D6BB6FC288C3FA892553A49049DA0911EFA7ADD00FFFC88612591DD046B4D0"
            && output.Report.UpdatedRecords.Count == 1 && body.Paragraphs[0].Text == "Native paragraph layout", "Paragraph edit differs from independent/native proof or immutable source.");
        Check(IWorkSourceDocument.Open(output.GetPackageBytes()), "Updated reset Native layout edit.");
        Check(Open("pages-paragraph-edited.pages"), "Native checked paragraph.");
        byte[] before = output.GetPackageBytes(); using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        try { edit.SetPagesBodyParagraphText(id, 1, body.Paragraphs[0].Text, cancellation.Token); throw new Exception("Paragraph edit ignored cancellation."); }
        catch (OperationCanceledException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Cancelled paragraph edit changed pending values."); }
        try { edit.SetPagesBodyParagraphText(id, 1, "Length change"); throw new Exception("Unsupported paragraph length was accepted."); }
        catch (NotSupportedException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Rejected paragraph edit changed pending values."); }
        foreach (int paragraph in new[] { 1, 3, 7 }) edit.SetPagesBodyParagraphText(id, paragraph, body.Paragraphs[paragraph - 1].Text);
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0 && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(), "Restoring text did not restore original native bytes.");
        Console.WriteLine("PASS | native Pages paragraph editing | three exact UTF-8 range edits, fixed UTF-16 offsets, preserved layout/unknown bytes, Apple edit/reopen, saved DOCX, cancellation, length rejection and restoration");
    }

    private static void Check(IWorkSourceDocument source, string finalParagraph) {
        var pages = source.ReadPages();
        Require(pages.HasEditableContent && pages.Body.IsComplete && pages.Paragraphs.Count == 8
            && pages.Paragraphs[0] == "Edited paragraph layout" && pages.Paragraphs[2] == "Start\tLeft\tCenter\tRight\t234.56"
            && pages.Paragraphs[5] == "Rich layout plain Zażółć 😀" && pages.Body.Paragraphs[5].Runs.Count == 2
            && pages.Paragraphs[6] == finalParagraph, "Edited Pages content or rich paragraph differs.");
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(source.BeginEdit().SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream(); converted.Value.Save(bytes); bytes.Position = 0;
        using var saved = WordDocument.Load(bytes);
        using var xml = WordprocessingDocument.Open(new MemoryStream(bytes.ToArray()), false);
        var paragraphs = xml.MainDocumentPart!.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
            .Select(item => string.Concat(item.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().SelectMany(run => run.ChildElements)
                .Select(node => node switch { DocumentFormat.OpenXml.Wordprocessing.TabChar => "\t",
                    DocumentFormat.OpenXml.Wordprocessing.Text text => text.Text, _ => "" })));
        Require(!saved.ValidateDocument().Any() && pages.Paragraphs.SequenceEqual(paragraphs), "Saved DOCX lost edited body content.");
    }
    private static IWorkSourceDocument Open(string name) => IWorkSourceDocument.Open(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name)));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
