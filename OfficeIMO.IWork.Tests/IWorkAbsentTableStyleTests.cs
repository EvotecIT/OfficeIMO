using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void SparseSourceBordersDoNotInheritDestinationGridOrFill(IWorkDocumentKind kind) {
        using var package = DimensionPackage(kind, tableStyleReference: ReferenceField(3, 40),
            styleRecords: [ArchiveRecord(40, 6003, BytesField(11, BytesField(51, BorderStroke(1))))]);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var converted = source.ToWordDocumentResult(policy);
            IWorkTable projected = Assert.Single(converted.Projection.Tables);
            Assert.True(projected.HasBorderDeclarations);
            Assert.Null(projected.GetBorders(1, 1).Top);
            Assert.NotNull(projected.GetBorders(1, 1).Right);
            converted.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved);
            WordTableCell cell = Assert.Single(reopened.Tables).Rows[0].Cells[0];
            Assert.Equal(WordBorderStyle.Nil, cell.Borders.TopStyle);
            Assert.Equal(WordBorderStyle.Nil, cell.Borders.LeftStyle);
            Assert.Equal(WordBorderStyle.Single, cell.Borders.RightStyle);
        } else {
            using var converted = source.ToPowerPointPresentationResult(policy);
            converted.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved);
            Assert.Empty(reopened.ValidateDocument());
            PowerPointTable table = Assert.Single(reopened.Slides[0].Tables);
            Assert.Null(table.StyleId);
            Assert.True(table.GetCell(0, 0).NoFill);
            saved.Position = 0;
            using var zip = new System.IO.Compression.ZipArchive(saved, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
            using var xml = zip.GetEntry("ppt/slides/slide1.xml")!.Open();
            var document = System.Xml.Linq.XDocument.Load(xml);
            System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var cell = document.Descendants(a + "tc").First();
            Assert.NotNull(cell.Element(a + "tcPr")!.Element(a + "lnT")!.Element(a + "noFill"));
        }
    }
}
