using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesNineLevelSmoke {
    internal static readonly string[] LevelNames = ["Root seven", "Level one", "Level two", "Level three",
        "Level four", "Level five", "Level six", "Level seven", "Level eight"];
    internal static readonly int[] Numbers = [7, 1, 1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 1, 1, 1, 1, 9, 3, 10,
        8, 1, 1, 1, 1, 1, 1, 1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1];
    internal static readonly int[] Levels = [0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 8, 4, 5, 6, 7, 8, 8, 0, 8,
        0, 1, 2, 3, 4, 5, 6, 7, 8, 0, 1, 2, 3, 4, 5, 6, 7, 8];

    internal static void Run() {
        byte[] created = Example().SaveBytes();
        Require(created.Length == 22332 && Hash(created) == "7DCAF3219B199CB33F89B37520642DCE96A10E4512BB81431232D079FA02C154"
            && created.SequenceEqual(Example().SaveBytes()), "Nine-level creation differs from the independent/native reference.");
        CheckWord(IWorkSourceDocument.Open(created), Numbers, Levels);
        byte[] original = Read("pages-nine-level-original.pages");
        Require(Hash(original) == "06032EC01C1BB8E03B4EDC9B3E7A189006D3CB9E142BFBE4459D048929DA6E9A", "Native nine-level original changed.");
        var source = IWorkSourceDocument.Open(original); var body = source.ReadPages().Body; var edit = source.BeginEdit();
        foreach (string text in LevelNames) {
            int number = body.Paragraphs.ToList().FindIndex(p => p.Text == text) + 1;
            edit.SetPagesBodyParagraphText(body.SourceIdentity!, number, text.ToUpperInvariant());
        }
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        var output = edit.SaveBytesResult(options); byte[] pending = output.GetPackageBytes();
        Require(Hash(pending) == "F104BE006883B9C741AA3B56E77BD826650A47392E10FD5B657BC5361222BC8B"
            && output.Report.UpdatedRecords.Count == 1 && body.Paragraphs[9].Text == "Level eight",
            "Nine-range preservation differs from independent/native proof or mutates the source.");
        var preserved = IWorkSourceDocument.Open(pending);
        Require(preserved.ReadPages().Body.Paragraphs.Skip(1).Take(9).Select(p => p.Text)
            .SequenceEqual(LevelNames.Select(text => text.ToUpperInvariant())), "A native level lost its text edit.");
        CheckWord(preserved, Numbers, Levels);
        byte[] native = Read("pages-nine-level-edited.pages");
        Require(Hash(native) == "B0BF86409D573682FF3D214AF543ABDE93F338AAB7ACA26A30F206593D878E52", "Native nine-level edit changed.");
        var nativeSource = IWorkSourceDocument.Open(native); var nativeBody = nativeSource.ReadPages().Body;
        Require(nativeBody.Paragraphs[1].Text == "Root novel" && nativeBody.Paragraphs[5].Text == "Level FOUR"
            && nativeBody.Paragraphs[9].Text == "Level EIGHT" && nativeBody.Paragraphs.Any(p => p.Text == "Native deep two"
                && p.ListLevel == 8 && p.ListNumber == 2), "Native edits/insertion did not survive reopen.");
        CheckWord(nativeSource, Numbers.Concat(new[] { 2 }).ToArray(), Levels.Concat(new[] { 8 }).ToArray());
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        try { edit.SetPagesBodyParagraphText(body.SourceIdentity!, 10, "Level eight", cancel.Token); throw new Exception("Nine-level edit ignored cancellation."); }
        catch (OperationCanceledException) { Require(pending.SequenceEqual(edit.SaveBytes(options)), "Cancellation changed pending edits."); }
        try { edit.SetPagesBodyParagraphText(body.SourceIdentity!, 10, "Short"); throw new Exception("Nine-level edit accepted a length change."); }
        catch (NotSupportedException) { Require(pending.SequenceEqual(edit.SaveBytes(options)), "Rejection changed pending edits."); }
        foreach (string text in LevelNames) {
            int number = body.Paragraphs.ToList().FindIndex(p => p.Text == text) + 1;
            edit.SetPagesBodyParagraphText(body.SourceIdentity!, number, text);
        }
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0 && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(),
            "Nine-level restoration did not restore original bytes.");
        Console.WriteLine("PASS | native Pages nine decimal levels | deterministic creation, ancestor resets/restarts, resumed scopes, nine exact text ranges, Apple edit/insertion/reopen and strict saved DOCX; tiered and broader text edits remain separate");
    }

    private static void CheckWord(IWorkSourceDocument source, int[] numbers, int[] levels) {
        var pages = source.ReadPages(); var items = pages.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        Require(pages.Body.IsFormattingComplete && pages.Diagnostics.Count == 0 && pages.SourceDeclarationIssues.Count == 0
            && pages.SourceReferenceIssues.Count == 0 && items.Select(p => p.ListNumber!.Value).SequenceEqual(numbers)
            && items.Select(p => p.ListLevel).SequenceEqual(levels), "Nine-level projection lost content or counters.");
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var saved = WordDocument.Load(bytes);
        Require(!saved.ValidateDocument().Any() && WordDocumentTraversal.BuildListIndices(saved).Values
            .Select(p => p.Index).SequenceEqual(numbers), "Saved nine-level DOCX counters changed.");
    }

    private static byte[] Read(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
    internal static IWorkPagesDocument Example(IWorkDecimalListMarkerStyle style = IWorkDecimalListMarkerStyle.Period) {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Nine decimal levels");
        var list = document.CreateNumberedList(style);
        for (int level = 0; level < 9; level++) list.AddItem(LevelNames[level], level, level == 0 ? 7 : null);
        list.AddItem("Deep continued two", 8); document.AddParagraph("Plain gap");
        list.AddItem("Deep after gap three", 8); list.AddItem("Middle restart four", 4, 4);
        foreach (int level in new[] { 5, 6, 7, 8 }) list.AddItem("Reset level " + level, level);
        list.AddItem("Deep restart nine", 8, 9);
        document.CreateNumberedList().AddItem("Other root three", restartAt: 3);
        list.AddItem("Resumed deep ten", 8); list.AddItem("Root eight");
        for (int level = 1; level < 9; level++) list.AddItem("Root reset level " + level, level);
        list.AddItem("Root restart four", restartAt: 4);
        for (int level = 1; level < 9; level++) list.AddItem("Restart level " + level, level);
        document.AddParagraph("End nine-level profile"); return document;
    }
}
