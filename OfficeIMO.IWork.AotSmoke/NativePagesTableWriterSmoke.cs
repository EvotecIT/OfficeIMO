using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesTableWriterSmoke {
    private const string AcceptedSha256 = "0869946A8507CCF60C528507E270B481B293A6184A4A357EA30D080303C355B3";

    internal static void Run() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Before 😀");
        var data = IWorkTableData.Create(3, 3).SetText(1, 1, "A😀B\r\nZażółć")
            .SetText(1, 2, "=SUM(A1:A2)").SetText(1, 3, "")
            .SetNumber(2, 1, -1234.5).SetBoolean(2, 2, true).SetBoolean(2, 3, false);
        document.AddTable(data, "Values", 108, 48, "Arial", 12, "003366");
        document.AddTable(IWorkTableData.Create(2, 2), "Empty", 100, 28);
        document.AddParagraph("After", new IWorkPagesParagraphLayout(1.5f, 6, 9));
        byte[] bytes = document.SaveBytes();
        Require(bytes.Length == 17_157 && Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedSha256,
            "NativeAOT Pages tables differ from the independently decoded Apple-accepted package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "Pages table creation is not deterministic.");
        var projection = IWorkSourceDocument.Open(bytes).ReadPages();
        Require(projection.Tables.Count == 2 && projection.Diagnostics.Count == 0
            && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0,
            "Pages table creation lost its supported graph.");
        var table = projection.Tables[0];
        Require(table.GetPadding(3, 3)?.BottomPoints == 4
            && table.GetVerticalAlignment(3, 3) == IWorkCellVerticalAlignment.Top
            && table.GetCell(1, 2)?.Kind == IWorkCellKind.Text
            && Equals(table.GetCell(2, 1)?.Value, -1234.5), "Pages effective layout or scalar types were lost.");
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Require(reopened.Tables.Count == 2 && reopened.Tables[0].Rows.Count == 3
            && reopened.Tables[1].Rows.Count == 2 && reopened.Tables[0].RowHeight.All(height => height?.ToTwips() == 880)
            && checked((short?)(reopened.Tables[0].Rows[2].Cells[2].MarginBottom?.ToTwips())) == 80
            && reopened.Tables[0].Rows[0].Cells[1].Paragraphs.Any(paragraph => paragraph.Text == "=SUM(A1:A2)")
            && !reopened.ValidateDocument().Any(), "Saved DOCX loses table separation, padding, height or literal text.");
        using var output = new MemoryStream(); output.WriteByte(9);
        byte[] before = output.ToArray();
        try {
            document.Save(output, new() { MaximumTableCells = 12 });
            throw new InvalidOperationException("Pages dense-grid bound was not enforced.");
        } catch (InvalidDataException) {
            Require(output.CanWrite && before.SequenceEqual(output.ToArray()), "Pages budget rejection changed its caller stream.");
        }
        var only = IWorkPagesDocument.Create(); only.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "Table only"));
        using var tableOnly = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(only.SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        tableOnly.RequireCompleteEditableReconstruction();
        Require(tableOnly.Value.Tables.Count == 1 && only.Paragraphs.Count == 0, "Table-only Pages body was not editable.");
        Console.WriteLine($"PASS | native Pages inline scalar-table creation | {AcceptedSha256} | {bytes.Length} bytes | deterministic body order, typed Unicode cells, inherited layout, strict saved/reopened DOCX, table-only body and atomic grid bounds");
        byte[] nativeBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-created-tables-edited.pages"));
        const string nativeHash = "F8A50C0312377A099C4C520B8DCDD3657391D9831C9C85D15474BFA740ADE2CD";
        Require(Convert.ToHexString(SHA256.HashData(nativeBytes)) == nativeHash, "Native Pages table edit fixture changed.");
        var native = IWorkSourceDocument.Open(nativeBytes);
        var nativeProjection = native.ReadPages();
        Require(nativeProjection.Tables.Count == 2 && nativeProjection.Diagnostics.Count == 0
            && nativeProjection.SourceDeclarationIssues.Count == 0 && nativeProjection.SourceReferenceIssues.Count == 0
            && Equals(nativeProjection.Tables[0].GetCell(2, 1)?.Value, 20.75)
            && nativeProjection.Tables[1].GetPadding(2, 2)?.BottomPoints == 4,
            "Native Pages edit or empty-grid region padding was lost.");
        using var nativeConverted = native.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        nativeConverted.RequireCompleteEditableReconstruction();
        using var nativeSaved = new MemoryStream(); nativeConverted.Value.Save(nativeSaved); nativeSaved.Position = 0;
        using var nativeReopened = WordDocument.Load(nativeSaved);
        Require(nativeReopened.Tables.Count == 2 && !nativeReopened.ValidateDocument().Any(), "Native edited Pages DOCX failed saved validation.");
        Console.WriteLine($"PASS | native Pages saved inline table edit | {nativeHash} | typed edit, empty-grid region padding and strict saved/reopened DOCX");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
