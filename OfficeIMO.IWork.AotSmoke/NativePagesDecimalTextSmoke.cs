using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesDecimalTextSmoke {
    internal static void Run() {
        var source = Open("pages-decimal-original.pages"); var body = source.ReadPages().Body;
        var edit = source.BeginEdit(); var id = body.SourceIdentity!;
        edit.SetPagesBodyParagraphText(id, 2, "Edited seven");
        edit.SetPagesBodyParagraphText(id, 3, "Child two");
        edit.SetPagesBodyParagraphText(id, 4, "Grandchild six");
        edit.SetPagesBodyParagraphText(id, 12, "Updated grandchild six");
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        var output = edit.SaveBytesResult(options);
        Require(Convert.ToHexString(SHA256.HashData(output.GetPackageBytes()))
            == "8E20D81E90F339CE7E7ACBC63D89A76F9B17CB8B656D9433F656865BE3D76AB7"
            && output.Report.UpdatedRecords.Count == 1 && body.Paragraphs[1].Text == "Parent seven",
            "Decimal-item edit differs from independent/native proof or immutable source.");
        Check(IWorkSourceDocument.Open(output.GetPackageBytes()), "Edited seven");
        var native = Open("pages-decimal-edited.pages");
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-decimal-edited.pages"))))
            == "4E9BE6834D5544DEC4605DA1027CF7E448974657BB66E4F15AFDD2798A58501F", "Native-resaved decimal input changed.");
        Check(native, "Native seven");
        byte[] before = output.GetPackageBytes(); using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        try { edit.SetPagesBodyParagraphText(id, 4, "Grandchild one", cancellation.Token); throw new Exception("Decimal edit ignored cancellation."); }
        catch (OperationCanceledException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Cancelled decimal edit changed pending values."); }
        try { edit.SetPagesBodyParagraphText(id, 4, "Short"); throw new Exception("Unsupported decimal length accepted."); }
        catch (NotSupportedException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Rejected decimal edit changed pending values."); }
        foreach (int paragraph in new[] { 2, 3, 4, 12 }) edit.SetPagesBodyParagraphText(id, paragraph, body.Paragraphs[paragraph - 1].Text);
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0 && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(),
            "Restoring decimal text did not restore original native bytes.");
        Console.WriteLine("PASS | native Pages decimal-item editing | four exact UTF-8 ranges, levels0..2, preserved numbering/indents/unknown bytes, Apple edit/reopen, saved DOCX, cancellation, length rejection and restoration");
    }

    private static void Check(IWorkSourceDocument source, string parent) {
        var pages = source.ReadPages(); int[] expected = [7, 1, 1, 2, 3, 2, 1, 5, 3, 6, 8, 1, 1, 4, 1, 1];
        Require(pages.HasEditableContent && pages.Body.IsFormattingComplete && pages.Diagnostics.Count == 0
            && pages.Paragraphs[1] == parent && pages.Paragraphs[2] == "Child two" && pages.Paragraphs[3] == "Grandchild six"
            && pages.Paragraphs[11] == "Updated grandchild six"
            && pages.Body.Paragraphs.Where(p => p.ListNumber.HasValue).Select(p => p.ListNumber!.Value).SequenceEqual(expected),
            "Edited decimal text or counters changed.");
        using var converted = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction(); using var bytes = new MemoryStream(); converted.Value.Save(bytes); bytes.Position = 0;
        using var saved = WordDocument.Load(bytes);
        Require(!saved.ValidateDocument().Any() && WordDocumentTraversal.BuildListIndices(saved).Values.Select(p => p.Index).SequenceEqual(expected),
            "Saved DOCX decimal counters changed.");
    }

    private static IWorkSourceDocument Open(string name) => IWorkSourceDocument.Open(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name)));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
