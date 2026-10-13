using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 1f)]
    [InlineData(IWorkDocumentKind.Numbers, 1f)]
    [InlineData(IWorkDocumentKind.Keynote, 1f)]
    [InlineData(IWorkDocumentKind.Pages, 0f)]
    [InlineData(IWorkDocumentKind.Numbers, 0f)]
    [InlineData(IWorkDocumentKind.Keynote, 0f)]
    public void Complete_explicit_grid_edges_shadow_unsupported_inherited_border_defaults(IWorkDocumentKind kind, float width) {
        using var package = FullySelectedBorderPackage(kind, width);
        var source = IWorkSourceDocument.Open(package);
        IWorkTable table;
        IReadOnlyList<IWorkDiagnostic> diagnostics;
        if (kind == IWorkDocumentKind.Pages) {
            var projection = source.ReadPages(); table = Assert.Single(projection.Tables); diagnostics = projection.Diagnostics;
        } else if (kind == IWorkDocumentKind.Numbers) {
            var projection = source.ReadNumbers(); table = Assert.Single(Assert.Single(projection.Sheets).Tables); diagnostics = projection.Diagnostics;
        } else {
            var projection = source.ReadKeynote(); table = Assert.Single(Assert.Single(projection.Slides).Tables); diagnostics = projection.Diagnostics;
        }
        Assert.DoesNotContain(diagnostics, d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
        Assert.True(table.HasBorderDeclarations);
        for (int row = 1; row <= 3; row++)
            for (int column = 1; column <= 2; column++) {
                var borders = table.GetBorders(row, column);
                foreach (var side in new[] { borders.Left, borders.Top, borders.Right, borders.Bottom })
                    AssertBorder(side, "000000", width);
            }
    }

    [Theory]
    [InlineData("horizontal-gap")]
    [InlineData("vertical-gap")]
    [InlineData("unsupported-run")]
    [InlineData("invalid-location")]
    [InlineData("equal-order-conflict")]
    public void Incomplete_or_rejected_explicit_grid_cannot_hide_unsupported_defaults(string defect) {
        using var package = FullySelectedBorderPackage(IWorkDocumentKind.Pages, 1, defect);
        var projection = IWorkSourceDocument.Open(package).ReadPages();
        Assert.Contains(projection.Diagnostics, d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
        Assert.NotEmpty(projection.SourceDeclarationIssues);
    }

    private static MemoryStream FullySelectedBorderPackage(IWorkDocumentKind kind, float width, string? defect = null) {
        // Native Pages DOCX imports retain rounded defaults while every active edge is explicitly replaced.
        byte[] unsupportedDefaults = Message(new[] { 58, 59, 60, 61 }
            .Select(field => BytesField(field, Message(BorderStroke(1), VarintField(4, 1)))).ToArray());
        var records = new List<byte[]> { ArchiveRecord(40, 6003, BytesField(11, unsupportedDefaults)) };
        var references = new List<byte[]>();
        ulong next = 61;
        foreach (bool horizontal in new[] { true, false }) {
            int lineCount = horizontal ? 3 : 2, length = horizontal ? 2 : 3;
            for (int line = 0; line <= lineCount; line++) {
                int field = horizontal ? line == lineCount ? 7 : 6 : line == lineCount ? 5 : 4;
                ulong index = (ulong)(line == lineCount ? line - 1 : line);
                int origin = line == 1 && (horizontal && defect == "horizontal-gap" || !horizontal && defect == "vertical-gap") ? 1 : 0;
                int runLength = length - origin;
                if (horizontal && line == 1 && defect == "invalid-location") { origin = length; runLength = 1; }
                byte[] stroke = BorderStroke(width, defect: horizontal && line == 1 && defect == "unsupported-run" ? "pattern" : null);
                AddLayer(field, index, origin, runLength, stroke);
                if (horizontal && line == 1 && defect == "equal-order-conflict")
                    AddLayer(field, index, 0, length, BorderStroke(2));
            }
        }
        records.Add(ArchiveRecord(60, 6305, Message(VarintField(1, 1), VarintField(2, 2), VarintField(3, 3), Message(references.ToArray()))));
        return DimensionPackage(kind, tableStyleReference: ReferenceField(3, 40), styleRecords: records.ToArray(),
            modelDeclarations: ReferenceField(49, 60));

        void AddLayer(int field, ulong index, int origin, int length, byte[] stroke) {
            ulong identifier = next++;
            references.Add(ReferenceField(field, identifier));
            records.Add(ArchiveRecord(identifier, 6306, Message(VarintField(1, index), BytesField(2,
                Message(VarintField(1, (ulong)origin), VarintField(2, (ulong)length), BytesField(3, stroke), VarintField(4, 1))))));
        }
    }
}
