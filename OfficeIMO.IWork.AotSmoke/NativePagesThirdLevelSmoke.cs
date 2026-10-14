using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesThirdLevelSmoke {
    internal static void Run() {
        byte[] bytes = Example().SaveBytes();
        Require(bytes.Length == 13763 && Convert.ToHexString(SHA256.HashData(bytes))
            == "50A5285D449DA02EEF7682A6A18B1F7103F6192304CC1F3173C52CAC5576F775",
            "Third-level creation differs from the independently decoded Apple-accepted input.");
        Require(bytes.SequenceEqual(Example().SaveBytes()), "Third-level creation is nondeterministic.");
        int[] expected = [7, 1, 1, 2, 3, 2, 1, 5, 3, 6, 8, 1, 1, 4, 1, 1];
        VerifySavedWord(bytes, expected);
        byte[] edited = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-third-level-edited.pages"));
        Require(Convert.ToHexString(SHA256.HashData(edited))
            == "B37A9B20F9BFED513498B70110D20ED8784B4B6513F17709DDDB557B7FB0C1C0", "Apple-edited third-level input changed.");
        VerifySavedWord(edited, expected.Concat(new[] { 2 }).ToArray());
        Console.WriteLine("PASS | native Pages third decimal level | Apple open/edit/save/reopen, ancestor resets, starts/restarts, scope resumption and strict saved DOCX; tiered numbering remains separate");
    }

    private static void VerifySavedWord(byte[] bytes, int[] expected) {
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Require(projection.Body.IsFormattingComplete && projection.Diagnostics.Count == 0
            && projection.SourceReferenceIssues.Count == 0 && projection.SourceDeclarationIssues.Count == 0,
            "Third-level source contains unqualified declarations.");
        Require(projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).Select(p => p.ListNumber!.Value)
            .SequenceEqual(expected), "Third-level counters changed.");
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var reopened = WordDocument.Load(output);
        Require(!reopened.ValidateDocument().Any() && WordDocumentTraversal.BuildListIndices(reopened).Values
            .Select(p => p.Index).SequenceEqual(expected), "Saved DOCX counters changed.");
    }

    private static IWorkPagesDocument Example() {

        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Three decimal levels");
        var list = document.CreateNumberedList();
        list.AddItem("Parent seven", restartAt: 7);
        list.AddItem("Child one", 1);
        list.AddItem("Grandchild one", 2); list.AddItem("Grandchild two", 2);
        document.AddParagraph("Plain gap");
        list.AddItem("Grandchild continued three", 2);
        list.AddItem("Child two resets grandchild", 1); list.AddItem("Grandchild reset one", 2);
        list.AddItem("Grandchild restart five", 2, 5);
        document.CreateNumberedList().AddItem("Separate root three", restartAt: 3);
        list.AddItem("Resumed grandchild six", 2);
        list.AddItem("Parent eight resets descendants");
        list.AddItem("Child reset one", 1); list.AddItem("Grandchild reset after root one", 2);
        list.AddItem("Parent restart four", restartAt: 4);
        list.AddItem("Child after restart one", 1); list.AddItem("Grandchild after restart one", 2);
        document.AddParagraph("End third-level profile");
        return document;
    }

    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
