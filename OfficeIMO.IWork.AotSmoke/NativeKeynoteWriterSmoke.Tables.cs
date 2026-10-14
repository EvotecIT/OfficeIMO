using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;

internal static partial class NativeKeynoteWriterSmoke {
    private const string TableAcceptedSha256 = "51D624971820BD161909E957DCB11ECE88B59F39F95268DB75DC166068FBCE88";

    private static void RunTables() {
        IWorkKeynoteDocument document = CreateTableExample();
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == TableAcceptedSha256,
            "NativeAOT table graph differs from the independently decoded and Apple-opened package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "NativeAOT table graph is nondeterministic.");
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(bytes).ReadKeynote();
        Require(projection.Slides.Count == 3 && projection.Slides.All(slide => slide.Tables.Count == 1),
            "NativeAOT table readback is incomplete.");
        IWorkTable table = projection.Slides[0].Tables.Single();
        Require(table.RowCount == 4 && table.ColumnCount == 4 && table.Name == "Values"
            && table.Cells.Single(cell => cell.Row == 2 && cell.Column == 3).Value is double number && number == 1200.5
            && table.Cells.Single(cell => cell.Row == 3 && cell.Column == 4).Value is false
            && table.Cells.Single(cell => cell.Row == 4 && cell.Column == 2).Value is string text && text == "=SUM(A1:A2)",
            "NativeAOT table dimensions or scalar types differ.");
        Require(projection.Slides[1].Drawables.Select(drawable => drawable.Kind).SequenceEqual(new[] {
            IWorkKeynoteDrawableKind.TextBox, IWorkKeynoteDrawableKind.Image,
            IWorkKeynoteDrawableKind.Table, IWorkKeynoteDrawableKind.TextBox
        }), "NativeAOT table stacking order differs.");
        Require(projection.Slides[2].Tables.Single().Cells.Count == 0, "NativeAOT empty table gained content.");
        Require(table.GetBorders(2, 2).Bottom is { Width: 1 } border && border.Color.ToRgbHex() == "000000",
            "NativeAOT table default borders differ.");
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(bytes),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var saved = PowerPointPresentation.Load(output);
        Require(saved.Slides.Count == 3 && !saved.ValidateDocument().Any()
            && saved.Slides[0].Tables.Single().GetCell(1, 2).Text == "1200.5"
            && saved.Slides[1].Tables.Single().GetCell(0, 0).Text == "A😀B\nZażółć — café",
            "NativeAOT table PPTX did not retain content through save/reopen.");
        foreach (var options in new[] { new IWorkKeynoteWriteOptions { MaximumTables = 2 },
            new IWorkKeynoteWriteOptions { MaximumTableCells = 20 } }) {
            using var unchanged = new MemoryStream(new byte[] { 9, 8, 7 });
            try { document.Save(unchanged, options); throw new InvalidOperationException("NativeAOT table budget was not enforced."); }
            catch (InvalidDataException) { }
            Require(unchanged.ToArray().SequenceEqual(new byte[] { 9, 8, 7 }), "Table-budget failure changed the destination.");
        }
        Console.WriteLine($"PASS | native Keynote tables | {TableAcceptedSha256} | {bytes.Length} bytes | Decimal128 numbers, text/booleans, empty grid, Unicode, stacking, deterministic creation, saved/reopened PPTX and table limits");
        RunNativeTableBorders();
    }

    private static void RunNativeTableBorders() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "native-table-borders.key"));
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == "2032CE6CDEE7F75DA8DE21D502B47D4E5C7F4EB4D152786A84DD81FEDA176D45",
            "NativeAOT controlled border source changed.");
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(bytes),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        var table = converted.Projection.Slides[0].Tables.Single();
        Require(table.GetBorders(2, 2).Bottom is { Width: 3 } red && red.Color.ToRgbHex() == "FF3333"
            && table.GetBorders(3, 2).Top is { Width: 3 } shared && shared.Color.ToRgbHex() == "FF3333"
            && table.GetBorders(2, 2).Right is { Width: 0 } && table.GetBorders(2, 3).Left is { Width: 0 }
            && table.GetBorders(1, 2).Top is { Width: 2 } && table.GetBorders(3, 3).Top is { Width: 0 },
            "NativeAOT border inheritance, visibility or ordered shared edges differ.");
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Require(!reopened.ValidateDocument().Any() && converted.Projection.SourceDeclarationIssues.Count == 0
            && converted.Projection.SourceReferenceIssues.Count == 0, "NativeAOT border PPTX or source qualification is incomplete.");
        Console.WriteLine("PASS | native Keynote table borders | native 15.4 source | inherited strokes, region visibility, ordered shared edits, disabled categories and saved/reopened PPTX");
    }

    private static IWorkKeynoteDocument CreateTableExample() {
        var document = IWorkKeynoteDocument.Create();
        var first = document.AddSlide("F3F7FC");
        first.AddText("Native editable table", 60, 30, 840, 60, "Arial", 32, "003366");
        var data = IWorkTableData.Create(4, 4)
            .SetText(1, 1, "Product").SetText(1, 2, "Units").SetText(1, 3, "Revenue").SetText(1, 4, "Active")
            .SetText(2, 1, "Widget").SetNumber(2, 2, 12).SetNumber(2, 3, 1200.5).SetBoolean(2, 4, true)
            .SetText(3, 1, "Gadget").SetNumber(3, 2, 0).SetNumber(3, 3, -25.75).SetBoolean(3, 4, false)
            .SetText(4, 1, "Zażółć 😀").SetText(4, 2, "=SUM(A1:A2)").SetText(4, 3, "");
        first.AddTable(data, 60, 110, 840, 300, "Values", "Arial", 24, "003366");
        first.AddText("Text, numeric and boolean cells; literal equals-sign text", 60, 440, 840, 60, "Arial", 22);
        var second = document.AddSlide();
        second.AddText("Multilingual text and mixed stacking order", 60, 30, 840, 60, "Arial", 28, "003366");
        second.AddImage(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "created-opaque.png")),
            100, 200, 160, 160, "Behind table");
        second.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "A😀B\r\nZażółć — café")
            .SetText(1, 2, "REPEATED").SetText(2, 1, "REPEATED").SetNumber(2, 2, 0.125),
            160, 130, 640, 300, "Unicode", "Times New Roman", 24);
        second.AddText("AFTER TABLE", 200, 460, 600, 60, "Arial", 28, "003366");
        var third = document.AddSlide();
        third.AddText("Empty editable table", 60, 30, 840, 60, "Arial", 32, "003366");
        third.AddTable(IWorkTableData.Create(1, 1), 60, 110, 600, 300, "Empty", fontSizePoints: 32);
        return document;
    }
}
