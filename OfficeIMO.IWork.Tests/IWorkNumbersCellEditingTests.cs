using System.Threading;
using OfficeIMO.Excel;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNumbersCellEditingTests {
    [Fact]
    public void Scalar_edits_compose_across_rows_and_tables_preserving_other_records_and_saved_Xlsx() {
        var source = Source(); var tables = Tables(source); var edit = source.BeginEdit();
        Apply(edit, tables);
        var result = edit.SaveBytesResult(AcceptCaches()); var saved = IWorkSourceDocument.Open(result.GetPackageBytes());
        Assert.Equal(File.ReadAllBytes(Fixture("numbers-preserved-scalars-v14.5.numbers")), result.GetPackageBytes());
        Assert.Equal(new ulong[] { 150, 175 }, result.Report.UpdatedRecords.Select(record => record.RecordIdentifier).Order());
        Assert.Equal(source.Records.Count, saved.Records.Count);
        foreach (var record in source.Records) {
            var output = saved.Records.Single(item => item.Identifier == record.Identifier && item.PayloadIndex == record.PayloadIndex);
            Assert.Equal(record.Versions, output.Versions); Assert.Equal(record.ObjectReferences, output.ObjectReferences);
            Assert.Equal(record.DataReferences, output.DataReferences);
            if (!result.Report.UpdatedRecords.Any(item => item.RecordIdentifier == record.Identifier))
                Assert.Equal(record.GetPayload(), output.GetPayload());
        }
        foreach (var entry in source.Entries.Where(entry => !result.Report.UpdatedEntryPaths.Contains(entry.Path)))
            Assert.Equal(entry.GetBytes(), saved.Entries.Single(item => item.Path == entry.Path).GetBytes());
        var data = Tables(saved).Single(table => table.Name == "Data");
        Assert.Equal(42.125, data.GetCell(2, 2)!.Value); Assert.Equal(false, data.GetCell(2, 3)!.Value);
        Assert.Equal(-7.5, data.GetCell(3, 2)!.Value); Assert.Equal(true, data.GetCell(3, 3)!.Value);
        Assert.Equal(20.75, tables[0].GetCell(2, 2)!.Value); // Immutable captured source.
        Assert.Equal(tables[0].GetCell(2, 2)!.NumberFormat!.DecimalPlaces, data.GetCell(2, 2)!.NumberFormat!.DecimalPlaces);
        Assert.Equal(tables[0].GetPadding(2, 2)!.LeftPoints, data.GetPadding(2, 2)!.LeftPoints);
        Assert.Equal(tables[0].GetPadding(2, 2)!.BottomPoints, data.GetPadding(2, 2)!.BottomPoints);
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(new MemoryStream(result.GetPackageBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(converted.Report.Diagnostics, item => item.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
        using var stream = new MemoryStream(); converted.Value.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        Assert.Equal(3, reopened.Sheets.Count); Assert.Empty(reopened.ValidateOpenXml());
        Assert.Equal(42.125, reopened.Sheets[0].CellAt(2, 2).GetValue<double>());
        Assert.False(reopened.Sheets[0].CellAt(2, 3).GetValue<bool>());
        Assert.Equal(123.75, reopened.Sheets[1].CellAt(1, 2).GetValue<double>());
        Assert.True(reopened.Sheets[1].CellAt(2, 1).GetValue<bool>());
    }

    [Fact]
    public void Rejected_or_cancelled_requests_preserve_pending_edits_and_original_values_restore_exact_source_bytes() {
        var source = Source(); var table = Tables(source)[0]; var edit = source.BeginEdit(); var id = table.SourceIdentity!;
        edit.SetNumbersCellNumber(id, 2, 2, 42.125); edit.SetNumbersCellBoolean(id, 2, 3, false);
        byte[] before = edit.SaveBytes(AcceptCaches());
        foreach (double value in new[] { double.NaN, double.NegativeInfinity, double.PositiveInfinity })
            Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetNumbersCellNumber(id, 2, 2, value));
        Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetNumbersCellNumber(id, 0, 2, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetNumbersCellBoolean(id, 2, table.ColumnCount + 1, false));
        Assert.Throws<NotSupportedException>(() => edit.SetNumbersCellNumber(id, 2, 3, 1));
        Assert.Throws<NotSupportedException>(() => edit.SetNumbersCellBoolean(id, 2, 2, true));
        Assert.Throws<NotSupportedException>(() => edit.SetNumbersCellNumber(id, 1, 1, 1));
        Assert.Throws<NotSupportedException>(() => edit.SetNumbersCellNumber(id, 4, 4, 1));
        Assert.Throws<ArgumentException>(() => edit.SetNumbersCellNumber(source.ReadNumbers().Sheets[0].SourceIdentity!, 2, 2, 1));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => edit.SetNumbersCellNumber(id, 2, 2, 1, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => edit.SetNumbersCellBoolean(id, 2, 3, true, cancellation.Token));
        Assert.Equal(before, edit.SaveBytes(AcceptCaches()));
        using var output = new MemoryStream(); output.WriteByte(9);
        Assert.Throws<InvalidOperationException>(() => edit.Save(output)); Assert.Equal(new byte[] { 9 }, output.ToArray());
        edit.SetNumbersCellNumber(id, 2, 2, (double)table.GetCell(2, 2)!.Value!);
        Assert.Single(edit.SaveBytesResult(AcceptCaches()).Report.UpdatedRecords);
        edit.SetNumbersCellBoolean(id, 2, 3, (bool)table.GetCell(2, 3)!.Value!);
        var restored = edit.SaveBytesResult(); Assert.Empty(restored.Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), restored.Report.OutputPackageContentHash);
    }

    [Fact]
    public void Formula_bearing_sources_are_rejected_before_any_cell_mutation() {
        var source = IWorkSourceDocument.Open(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser", "test-10-formulas.numbers"));
        var tables = Tables(source); Assert.Contains(tables.SelectMany(table => table.Cells), cell => cell.SourceFormulaIsDeclared);
        var edit = source.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetNumbersCellNumber(tables[0].SourceIdentity!, 1, 1, 1));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void A_populated_legacy_representation_is_rejected_even_when_modern_projection_is_complete() {
        var original = Source(); var tile = original.Records.Single(record => record.Identifier == 150 && record.PayloadIndex == 0);
        byte[] payload = tile.GetPayload();
        // The native fixture's legacy B2 record starts at12; its empty type byte is at13.
        IWorkPreservingWireEdit.ReplaceBytesRange(payload, new[] { (5, 1), (3, -1) }, 13, new byte[] { 2 }, new(), default);
        var changed = IWorkPreservationPackageWriter.Write(original, new Dictionary<IWorkArchiveRecord, byte[]> { [tile] = payload }, AcceptCaches(), default);
        var source = IWorkSourceDocument.Open(changed.GetPackageBytes()); Assert.True(source.ReadNumbers().HasEditableContent);
        var edit = source.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetNumbersCellNumber(Tables(source)[0].SourceIdentity!, 2, 2, 42));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Qualified_edited_sources_retain_typed_values_styles_and_all_tables(bool nativeEdited) {
        string name = "numbers-preserved-scalars" + (nativeEdited ? "-edited" : "") + "-v14.5.numbers";
        var source = IWorkSourceDocument.Open(Fixture(name)); var projection = source.ReadNumbers();
        Assert.True(projection.HasEditableContent); Assert.Empty(projection.Diagnostics);
        Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var tables = Tables(source); Assert.Equal(2, projection.Sheets.Count); Assert.Equal(3, tables.Length);
        Assert.Equal(nativeEdited ? 43.125 : 42.125, tables[0].GetCell(2, 2)!.Value);
        Assert.Equal(false, tables[0].GetCell(2, 3)!.Value); Assert.Equal(-7.5, tables[0].GetCell(3, 2)!.Value);
        Assert.Equal(true, tables[0].GetCell(3, 3)!.Value); Assert.Equal("Zażółć 😀", tables[0].GetCell(2, 4)!.Value);
        Assert.Equal(123.75, tables[1].GetCell(1, 2)!.Value); Assert.Equal(true, tables[1].GetCell(2, 1)!.Value);
        Assert.Empty(tables[2].Cells); Assert.Equal(2, tables[2].RowCount); Assert.Equal(3, tables[2].ColumnCount);
        // A second edit session over Apple's resave remains usable and restores the original encoding.
        var edit = source.BeginEdit(); edit.SetNumbersCellNumber(tables[0].SourceIdentity!, 2, 2, 44.125);
        edit.SetNumbersCellNumber(tables[0].SourceIdentity!, 2, 2, (double)tables[0].GetCell(2, 2)!.Value!);
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Theory]
    [InlineData(1.2345678901234567)]
    [InlineData(double.MaxValue)]
    [InlineData(double.Epsilon)]
    public void Finite_decimal_values_round_trip_without_changing_the_original_scalar_on_reset(double value) {
        var model = IWorkNumbersDocument.Create();
        model.AddSheet("Numbers").AddTable(IWorkTableData.Create(1, 1).SetNumber(1, 1, 0.125), "Values");
        var source = IWorkSourceDocument.Open(model.SaveBytes()); var table = Tables(source)[0]; var edit = source.BeginEdit();
        edit.SetNumbersCellNumber(table.SourceIdentity!, 1, 1, value);
        var saved = IWorkSourceDocument.Open(edit.SaveBytes(AcceptCaches()));
        Assert.Equal(value, Tables(saved)[0].GetCell(1, 1)!.Value);
        edit.SetNumbersCellNumber(table.SourceIdentity!, 1, 1, 0.125);
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    internal static void Apply(IWorkNativeEditSession edit, IReadOnlyList<IWorkTable> tables) {
        var data = tables.Single(table => table.Name == "Data").SourceIdentity!;
        var totals = tables.Single(table => table.Name == "Totals").SourceIdentity!;
        edit.SetNumbersCellNumber(data, 2, 2, 42.125); edit.SetNumbersCellBoolean(data, 2, 3, false);
        edit.SetNumbersCellNumber(data, 3, 2, -7.5); edit.SetNumbersCellBoolean(data, 3, 3, true);
        edit.SetNumbersCellNumber(totals, 1, 2, 123.75); edit.SetNumbersCellBoolean(totals, 2, 1, true);
    }

    private static IWorkSourceDocument Source() => IWorkSourceDocument.Open(Fixture("numbers-created-edited-v14.5.numbers"));
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static IWorkTable[] Tables(IWorkSourceDocument source) => source.ReadNumbers().Sheets.SelectMany(sheet => sheet.Tables).ToArray();
    private static IWorkNativeSaveOptions AcceptCaches() => new() { AllowUnassessedVisualCaches = true };
}
