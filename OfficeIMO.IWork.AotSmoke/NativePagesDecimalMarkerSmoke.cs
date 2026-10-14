using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesDecimalMarkerSmoke {
    internal static void Run() {
        Check(IWorkDecimalListMarkerStyle.Parentheses, "parentheses",
            "8E77AB2660E2B4A8742A70AB450B7153FD677449EFB13E53173A53781355A04C",
            "EC15D837258F8F5D30BC278A8BEE0E58A700DBB582C8586EB93AE4D0407457B9",
            "49AFF287CDF8F4B6673554683F767996C85FD8478E07B299521D08EE199E8919",
            "CE4B42947D6E111CC54D55A65FAB7B658A96CA32C2D62CFA828731580D0263E4");
        Check(IWorkDecimalListMarkerStyle.ClosingParenthesis, "closing",
            "E14BF8D07C445EAA2139349ABF1FA9F17708AC842FD12AA5F6F0731476FB0AA4",
            "39620F2CB28B0A3C0A9CCEAA22605F5F2CD293CD9B1C2D604E010A3EF1CB1B59",
            "213672CFD210690EA9749C5135A2D64017E2CFF901C1E7D6BAA689BCCEFB5F99",
            "8915D9303B604B196072C0BE41DACC840AE363D4CD03FD16FAF3F1F7B3DA0765");
        Console.WriteLine("PASS | native Pages decimal marker forms | deterministic (1)/1) creation, nine-level counters, exact text preservation, Apple edit/insertion/reopen and strict saved DOCX markers; tiered and mixed ancestor forms remain partial");
    }

    private static void Check(IWorkDecimalListMarkerStyle style, string name,
        string createdHash, string originalHash, string preservedHash, string nativeHash) {
        byte[] created = NativePagesNineLevelSmoke.Example(style).SaveBytes();
        Require(created.Length == 22332 && Hash(created) == createdHash
            && created.SequenceEqual(NativePagesNineLevelSmoke.Example(style).SaveBytes()), "Decimal marker creation changed.");
        CheckWord(IWorkSourceDocument.Open(created), style);
        byte[] original = Read(name, "original"); Require(Hash(original) == originalHash, "Apple marker input changed.");
        var source = IWorkSourceDocument.Open(original); var body = source.ReadPages().Body;
        CheckWord(source, style);
        var edit = source.BeginEdit();
        foreach (string text in NativePagesNineLevelSmoke.LevelNames) {
            int number = body.Paragraphs.ToList().FindIndex(p => p.Text == text) + 1;
            edit.SetPagesBodyParagraphText(body.SourceIdentity!, number, text.ToUpperInvariant());
        }
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        var output = edit.SaveBytesResult(options); byte[] pending = output.GetPackageBytes();
        Require(Hash(pending) == preservedHash && output.Report.UpdatedRecords.Count == 1
            && output.Report.UpdatedEntryPaths.Count == 1 && body.Paragraphs[1].Text == "Root seven",
            "Decimal marker text patch differs from independent/native evidence or mutates input.");
        var preserved = IWorkSourceDocument.Open(pending);
        Require(preserved.ReadPages().Body.Paragraphs.Skip(1).Take(9).Select(p => p.Text)
            .SequenceEqual(NativePagesNineLevelSmoke.LevelNames.Select(text => text.ToUpperInvariant())), "A marker item lost its edit.");
        CheckWord(preserved, style);
        byte[] native = Read(name, "edited"); Require(Hash(native) == nativeHash, "Apple marker edit changed.");
        var nativeSource = IWorkSourceDocument.Open(native); var nativeBody = nativeSource.ReadPages().Body;
        Require(nativeBody.Paragraphs[1].Text == "Root novel" && nativeBody.Paragraphs[9].Text == "Level EIGHT"
            && nativeBody.Paragraphs.Any(p => p.Text == "Native deep two" && p.ListLevel == 8 && p.ListNumber == 2),
            "Native marker edit/insertion did not survive reopen.");
        CheckWord(nativeSource, style, edited: true);
        int deepest = body.Paragraphs.ToList().FindIndex(p => p.Text == "Level eight") + 1;
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        try { edit.SetPagesBodyParagraphText(body.SourceIdentity!, deepest, "Level eight", cancel.Token); throw new Exception("Marker edit ignored cancellation."); }
        catch (OperationCanceledException) { Require(pending.SequenceEqual(edit.SaveBytes(options)), "Cancellation changed marker edits."); }
        try { edit.SetPagesBodyParagraphText(body.SourceIdentity!, deepest, "short"); throw new Exception("Marker edit accepted a length change."); }
        catch (NotSupportedException) { Require(pending.SequenceEqual(edit.SaveBytes(options)), "Rejection changed marker edits."); }
        foreach (string text in NativePagesNineLevelSmoke.LevelNames) {
            int number = body.Paragraphs.ToList().FindIndex(p => p.Text == text) + 1;
            edit.SetPagesBodyParagraphText(body.SourceIdentity!, number, text);
        }
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0 && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(),
            "Restoring marker text did not restore the original bytes.");
    }

    private static void CheckWord(IWorkSourceDocument source, IWorkDecimalListMarkerStyle style, bool edited = false) {
        var pages = source.ReadPages(); var items = pages.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        int[] numbers = edited ? NativePagesNineLevelSmoke.Numbers.Concat(new[] { 2 }).ToArray() : NativePagesNineLevelSmoke.Numbers;
        int[] levels = edited ? NativePagesNineLevelSmoke.Levels.Concat(new[] { 8 }).ToArray() : NativePagesNineLevelSmoke.Levels;
        string sample = style == IWorkDecimalListMarkerStyle.Parentheses ? "(1)" : "1)";
        Require(pages.Body.IsFormattingComplete && pages.Diagnostics.Count == 0 && pages.SourceDeclarationIssues.Count == 0
            && pages.SourceReferenceIssues.Count == 0 && items.Select(p => p.ListNumber!.Value).SequenceEqual(numbers)
            && items.Select(p => p.ListLevel).SequenceEqual(levels)
            && items.All(p => p.ListLabel == (p.Text == "Other root three" ? "1." : sample)), "Marker projection lost counters or punctuation.");
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var saved = WordDocument.Load(bytes);
        string[] markers = items.Select(p => p.Text == "Other root three" ? "3."
            : style == IWorkDecimalListMarkerStyle.Parentheses ? "(" + p.ListNumber + ")" : p.ListNumber + ")").ToArray();
        Require(!saved.ValidateDocument().Any() && WordDocumentTraversal.BuildListMarkers(saved).Values.Select(p => p.Marker).SequenceEqual(markers)
            && WordDocumentTraversal.BuildListIndices(saved).Values.Select(p => p.Index).SequenceEqual(numbers), "Saved DOCX markers changed.");
    }

    private static byte[] Read(string name, string suffix) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-decimal-" + name + "-" + suffix + ".pages"));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
