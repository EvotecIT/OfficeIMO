using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableDecorationEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_style_edits_match_independent_producer_bytes_and_restore_exact_source(bool reverse) {
        byte[] source = Fixture(reverse), expected = Fixture(!reverse);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable snapshot = document.Pages[0].Tables.Single();
        Apply(document, !reverse);
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved);
        Assert.Equal(Art(expected), Art(saved));
        PublisherTable reopened = PublisherDocument.Load(saved).Pages[0].Tables.Single();
        PublisherTable native = PublisherDocument.Load(expected).Pages[0].Tables.Single();
        Assert.Equal(native.Cells.Select(cell => cell.FillColor), reopened.Cells.Select(cell => cell.FillColor));
        Assert.Equal(native.BorderSegments, reopened.BorderSegments);
        Assert.Equal(snapshot.Cells.Select(cell => cell.Text), reopened.Cells.Select(cell => cell.Text));
        Assert.Equal(reverse ? OfficeColor.FromRgb(192, 224, 192) : OfficeColor.FromRgb(255, 224, 160), snapshot.Cells[0].FillColor);
        Assert.Equal(reverse ? 1D : 2D, snapshot.BorderSegments[0].Side.Width);
        Apply(document, reverse);
        Assert.Equal(source, document.ToPubBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Cell_fill_visibility_saves_and_restores_to_exact_original_source(bool insert) {
        byte[] source = Fixture(false);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables[0];
        int row = insert ? 1 : 0;
        OfficeColor? prior = original.Cells.Single(cell => cell.RowIndex == row && cell.ColumnIndex == 0).FillColor;
        OfficeColor? requested = insert ? OfficeColor.FromRgb(20, 40, 60) : null;
        document.SetTableCellFillColor(original.Id, row, 0, requested);
        Assert.Equal(requested, document.Pages[0].Tables[0].Cells.Single(cell => cell.RowIndex == row && cell.ColumnIndex == 0).FillColor);
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value);
        Assert.Equal(requested, PublisherDocument.Load(result.Value).Pages[0].Tables[0].Cells.Single(cell => cell.RowIndex == row && cell.ColumnIndex == 0).FillColor);
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_TABLE_FILL_UNQUALIFIED" && item.LossKind == OfficeConversionLossKind.Unassessed);
        document.SetTableCellFillColor(original.Id, row, 0, prior);
        Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Nonrepresentable_border_width_refuses_native_output_without_rounding_the_model() {
        byte[] source = Fixture(false);
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = document.Pages[0].Tables[0].Id;
        document.SetTableBorderStyle(id, 0, 0, 1, 0, new OfficeBorderSide(OfficeColor.Black, 1D / 3));
        Assert.Equal(1D / 3, document.Pages[0].Tables[0].BorderSegments[0].Side.Width);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        document.SetTableBorderStyle(id, 0, 0, 1, 0, new OfficeBorderSide(OfficeColor.FromRgb(255, 0, 0), 2));
        Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Authored_fills_follow_merge_anchors_and_invalid_or_cancelled_edits_are_atomic() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        uint id = document.AddTable(document.Pages[0].Id, 72, 144, new[] { 120D, 180D }, new[] { 42D, 54D }).Id;
        OfficeColor anchor = OfficeColor.FromRgb(20, 40, 60), covered = OfficeColor.FromRgb(80, 100, 120);
        document.SetTableCellFillColor(id, 0, 0, anchor);
        document.SetTableCellFillColor(id, 0, 1, covered);
        PublisherTable snapshot = document.Pages[0].Tables[0];
        document.MergeTableCells(id, 0, 0, 1, 2);
        Assert.Equal(anchor, document.Pages[0].Tables[0].Cells[0].FillColor);
        Assert.Equal(covered, snapshot.Cells[1].FillColor);
        string before = document.ToSvg();
        Assert.Throws<ArgumentException>(() => document.SetTableCellFillColor(id, 0, 1, anchor));
        Assert.Throws<ArgumentException>(() => document.SetTableCellFillColor(id, 0, 0, new OfficeColor(1, 2, 3, 128)));
        Assert.Throws<OperationCanceledException>(() => document.SetTableCellFillColor(id, 0, 0, null, new CancellationToken(true)));
        Assert.Throws<ArgumentException>(() => document.SetTableBorderStyle(id, 0, 0, 1, 1, new OfficeBorderSide(anchor, 1)));
        Assert.Equal(before, document.ToSvg());
        document.SetTableCellFillColor(id, 0, 0, null);
        Assert.Null(document.Pages[0].Tables[0].Cells[0].FillColor);
    }

    [Fact]
    public void Invalid_shared_border_edits_retain_all_styles_and_source_bytes() {
        byte[] source = Fixture(false);
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = document.Pages[0].Tables[0].Id;
        Assert.Throws<ArgumentException>(() => document.SetTableBorderStyle(id, 0, 0, 1, 0, new OfficeBorderSide(OfficeColor.Black, 0)));
        Assert.Throws<ArgumentException>(() => document.SetTableBorderStyle(id, 0, 0, 1, 0, new OfficeBorderSide(OfficeColor.Black, 1, OfficeStrokeDashStyle.Dash)));
        Assert.Throws<ArgumentException>(() => document.SetTableBorderStyle(id, 0, 0, 1, 1, new OfficeBorderSide(OfficeColor.Black, 1)));
        Assert.Throws<ArgumentException>(() => document.SetTableBorderStyle(id, 0, 0, 3, 0, new OfficeBorderSide(OfficeColor.Black, 1)));
        Assert.Throws<OperationCanceledException>(() => document.SetTableBorderStyle(id, 0, 0, 1, 0, new OfficeBorderSide(OfficeColor.Black, 1), new CancellationToken(true)));
        Assert.Equal(source, document.ToPubBytes());
    }

    private static void Apply(PublisherDocument document, bool edited) {
        uint id = document.Pages[0].Tables[0].Id;
        document.SetTableCellFillColor(id, 0, 0, edited ? OfficeColor.FromRgb(192, 224, 192) : OfficeColor.FromRgb(255, 224, 160));
        document.SetTableBorderStyle(id, 0, 0, 1, 0, new OfficeBorderSide(edited ? OfficeColor.FromRgb(0, 128, 0) : OfficeColor.FromRgb(255, 0, 0), edited ? 1 : 2));
        document.SetTableBorderStyle(id, 1, 0, 1, 1, new OfficeBorderSide(OfficeColor.FromRgb(128, 0, 128), edited ? 3 : 1));
    }
    private static byte[] Fixture(bool edited) => File.ReadAllBytes(PublisherNativeTests.Fixture(edited
        ? "NativeComparison/TableDecorationEdited/native-tabledecorationedited.pub"
        : "NativeComparison/TableDecoration/native-tabledecoration.pub"));
    private static byte[] Art(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? file, out string? error), error);
        return file!.Streams["Escher/EscherStm"];
    }
}
