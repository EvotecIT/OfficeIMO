using OfficeIMO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableBorderAuthoringTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Shared_grid_ranges_create_split_replace_and_clear_without_changing_other_edges(bool vertical) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        uint id = document.AddTable(document.Pages[0].Id, 72, 144, new[] { 50D, 60D, 70D, 80D }, new[] { 40D, 50D, 60D, 70D }).Id;
        var blue = new OfficeBorderSide(OfficeColor.FromRgb(0, 0, 255), 2);
        var red = new OfficeBorderSide(OfficeColor.FromRgb(255, 0, 0), 3);
        // One perpendicular edge and a parallel edge on another grid line.
        document.SetTableBorderStyle(id, 0, 0, vertical ? 0 : 4, vertical ? 4 : 0, blue);
        document.SetTableBorderStyle(id, vertical ? 0 : 2, vertical ? 2 : 0, vertical ? 4 : 2, vertical ? 2 : 4, blue);
        Set(0, 4, blue);
        PublisherTable snapshot = document.Pages[0].Tables[0];
        Set(1, 3, red);
        Assert.Equal(new[] { (0, 1, blue), (1, 3, red), (3, 4, blue) }, Selected());
        Assert.Equal(3, snapshot.BorderSegments.Count);
        Set(2, 4, null);
        Assert.Equal(new[] { (0, 1, blue), (1, 2, red) }, Selected());
        Assert.Equal(4, document.Pages[0].Tables[0].BorderSegments.Count);
        Set(0, 4, red);
        Assert.Equal(new[] { (0, 4, red) }, Selected());
        Set(0, 4, null);
        Assert.Empty(Selected());
        Assert.Equal(snapshot.BorderSegments.Take(2), document.Pages[0].Tables[0].BorderSegments);
        string empty = document.ToSvg();
        Set(0, 4, null);
        Assert.Equal(empty, document.ToSvg());

        void Set(int from, int to, OfficeBorderSide? side) => document.SetTableBorderStyle(id,
            vertical ? from : 0, vertical ? 0 : from, vertical ? to : 0, vertical ? 0 : to, side);
        (int, int, OfficeBorderSide)[] Selected() => document.Pages[0].Tables[0].BorderSegments
            .Where(edge => vertical ? edge.FirstColumn == 0 && edge.LastColumn == 0 : edge.FirstRow == 0 && edge.LastRow == 0)
            .Select(edge => (vertical ? edge.FirstRow : edge.FirstColumn, vertical ? edge.LastRow : edge.LastColumn, edge.Side)).ToArray();
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Native_border_record_add_remove_and_range_changes_refuse_output_even_with_allow(int operation) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableDecoration/native-tabledecoration.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = document.Pages[0].Tables[0].Id;
        // Clearing an already absent edge does not change native representation.
        document.SetTableBorderStyle(id, 2, 0, 2, 1, null);
        Assert.Equal(source, document.ToPubBytes());
        var black = new OfficeBorderSide(OfficeColor.Black, 1);
        if (operation == 0) document.SetTableBorderStyle(id, 1, 1, 2, 1, black);
        else if (operation == 1) document.SetTableBorderStyle(id, 0, 0, 1, 0, null);
        else document.SetTableBorderStyle(id, 0, 0, 0, 2, black);
        Assert.Equal(operation == 0 ? 7 : operation == 1 ? 5 : 6, document.Pages[0].Tables[0].BorderSegments.Count);
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    [Fact]
    public void Merged_interiors_and_merging_a_bordered_table_refuse_without_publishing_partial_state() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        uint id = document.AddTable(document.Pages[0].Id, 72, 144, new[] { 120D, 180D }, new[] { 42D, 54D }).Id;
        document.MergeTableCells(id, 0, 0, 1, 2);
        string before = document.ToSvg();
        Assert.Throws<ArgumentException>(() => document.SetTableBorderStyle(id, 0, 1, 1, 1, new OfficeBorderSide(OfficeColor.Black, 1)));
        Assert.Equal(before, document.ToSvg());
        document.SetTableBorderStyle(id, 0, 0, 0, 2, new OfficeBorderSide(OfficeColor.Black, 1));
        before = document.ToSvg();
        document.MergeTableCells(id, 0, 0, 1, 2); // Existing span is unchanged.
        Assert.Throws<NotSupportedException>(() => document.MergeTableCells(id, 0, 0, 2, 2));
        Assert.Equal(before, document.ToSvg());
        Assert.Equal(3, document.Pages[0].Tables[0].Cells.Count);
    }

    [Fact]
    public void Border_creation_obeys_the_copied_projection_budget_and_cancels_atomically() {
        var options = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 6 } };
        PublisherDocument document = PublisherDocument.Create(612, 792, options);
        uint id = document.AddTable(document.Pages[0].Id, 72, 144, new[] { 120D, 180D }, new[] { 42D }).Id;
        options.Limits.MaxItems = 100;
        var side = new OfficeBorderSide(OfficeColor.Black, 1);
        var ranges = new[] { (0,0,0,1), (0,1,0,2), (1,0,1,1), (1,1,1,2), (0,0,1,0), (0,1,1,1), (0,2,1,2) };
        bool refused = false;
        foreach (var range in ranges) {
            string snapshot = document.ToSvg();
            int count = document.Pages[0].Tables[0].BorderSegments.Count;
            try { document.SetTableBorderStyle(id, range.Item1, range.Item2, range.Item3, range.Item4, side); }
            catch (InvalidDataException) {
                Assert.Equal(snapshot, document.ToSvg());
                Assert.Equal(count, document.Pages[0].Tables[0].BorderSegments.Count);
                refused = true; break;
            }
        }
        Assert.True(refused);
        Assert.InRange(document.Pages[0].Tables[0].BorderSegments.Count, 1, ranges.Length - 1);
        string before = document.ToSvg();
        Assert.Throws<OperationCanceledException>(() => document.SetTableBorderStyle(id, 0, 0, 0, 1, null, new CancellationToken(true)));
        Assert.Equal(before, document.ToSvg());
    }
}
