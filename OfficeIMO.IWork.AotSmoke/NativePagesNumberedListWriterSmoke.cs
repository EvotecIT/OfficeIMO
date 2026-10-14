using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesNumberedListWriterSmoke {
    private const string AcceptedHash = "EC6CA0C63892A5D266478BE084A585F25721F269723B861BC5F163354A91606D";
    private const string EditedHash = "4A7C98DE6F2B0F95AD04F0591D36788D1CDAD11CCA24222A828D5D3ABFA0022A";

    internal static void Run() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Owned decimal lists 😀");
        var list = document.CreateNumberedList(); list.AddItem("Parent seven", restartAt: 7); list.AddItem("Parent eight");
        document.AddParagraph("Plain gap"); list.AddItem("Continued nine"); list.AddItem("Child one", 1); list.AddItem("Child two", 1);
        document.AddTable(IWorkTableData.Create(1, 1).SetText(1, 1, "Between list items"), "Values", 180, 28);
        list.AddItem("Parent ten"); list.AddItem("Reset child one", 1); list.AddItem("Child restart five", 1, 5);
        document.CreateNumberedList().AddItem("Separate parent three", restartAt: 3);
        document.AddImage(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "created-alpha.png")), 72, 48, "Between lists");
        list.AddItem("Resumed child six", 1); list.AddItem("Resumed parent eleven"); list.AddItem("Reset child one again", 1);
        list.AddItem("Parent restart four", restartAt: 4); list.AddItem("Parent five", layout: new(1.25f, 3, 6));
        byte[] bytes = document.SaveBytes();
        Require(bytes.Length == 17_340 && Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedHash,
            "NativeAOT Pages lists differ from the independently decoded Apple-accepted package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "Native Pages list creation is nondeterministic.");
        VerifySavedWord(bytes, [7, 8, 9, 1, 2, 10, 1, 5, 3, 6, 11, 1, 4, 5]);
        using var unchanged = new MemoryStream(); unchanged.WriteByte(9);
        try {
            document.Save(unchanged, new() { MaximumNumberedLists = 1 });
            throw new InvalidOperationException("The Pages list budget was not enforced.");
        } catch (InvalidDataException) { }
        Require(unchanged.CanWrite && unchanged.ToArray().SequenceEqual(new byte[] { 9 }), "List rejection changed the caller stream.");
        var only = IWorkPagesDocument.Create(); var onlyList = only.CreateNumberedList();
        onlyList.AddItem("Only parent"); onlyList.AddItem("Only child", 1);
        Require(Convert.ToHexString(SHA256.HashData(only.SaveBytes())) == "C6BA923DA791EB24BCA60D58454CEBDDD6D52B84CA7AE7EA25EF638C7A2D9D26",
            "The Apple-accepted list-only profile changed.");
        using var onlyWord = IWorkSourceDocument.Open(only.SaveBytes()).ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        onlyWord.RequireCompleteEditableReconstruction();
        foreach (string separator in new[] { "\u2028", "\u2029" }) {
            try {
                onlyList.AddItem("First" + separator + "Second", restartAt: 100);
                throw new InvalidOperationException("A Unicode separator was accepted in a single list item.");
            } catch (ArgumentException) { }
        }
        Require(onlyList.Items.Count == 2 && only.Paragraphs.Count == 2 && onlyList.AddItem("Next parent").ListNumber == 2,
            "Rejected list separators changed the document or its counters.");
        Console.WriteLine($"PASS | native Pages decimal-list creation | {AcceptedHash} | deterministic groups, two levels, starts/restarts, body/table/image continuation, strict saved DOCX and bounds");
        byte[] edited = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-created-numbering-edited.pages"));
        Require(Convert.ToHexString(SHA256.HashData(edited)) == EditedHash, "Apple-edited Pages list fixture changed.");
        VerifySavedWord(edited, [7, 8, 9, 10, 1, 2, 11, 1, 5, 3, 6, 11, 1, 4, 5]);
        Console.WriteLine($"PASS | native Pages edited decimal lists | {EditedHash} | Apple insert/save/reopen, sparse restart flags and strict saved DOCX");
    }

    private static void VerifySavedWord(byte[] bytes, int[] expected) {
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Require(projection.Diagnostics.Count == 0 && projection.SourceDeclarationIssues.Count == 0
            && projection.SourceReferenceIssues.Count == 0 && projection.Body.IsFormattingComplete,
            "Pages lists contain unqualified source declarations.");
        Require(projection.Body.Paragraphs.Where(item => item.ListNumber.HasValue)
            .Select(item => item.ListNumber!.Value).SequenceEqual(expected), "Pages list projection counters changed.");
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordDocument.Load(output);
        Require(WordDocumentTraversal.BuildListIndices(saved).Values.Select(item => item.Index).SequenceEqual(expected)
            && !saved.ValidateDocument().Any() && saved.Images.Count == 1 && saved.Tables.Count == 1,
            "Saved DOCX lost list counters or body objects.");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
