using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRawStructureTests {
    [Fact]
    public void RawStructure_ProjectsActiveObjectsWithoutExposingMutableParserObjectsOrStreamBytes() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Raw structure"))
            .ToBytes();

        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        PdfRawObjectView catalog = Assert.IsType<PdfRawObjectView>(raw.GetObject(raw.CatalogObjectNumber!.Value));

        Assert.Equal(raw.TotalObjectCount, raw.Objects.Count);
        Assert.Equal(PdfRawValueKind.Dictionary, catalog.Value.Kind);
        Assert.Equal("Catalog", catalog.Value.Entries["Type"].Text);
        Assert.NotEmpty(raw.Revisions);
        Assert.Contains("/Root", raw.TrailerPreview, StringComparison.Ordinal);
        Assert.Contains(raw.Objects, obj => obj.Value.Kind == PdfRawValueKind.Stream && obj.Value.StreamLength > 0);
        Assert.Throws<NotSupportedException>(() =>
            ((IDictionary<string, PdfRawValue>)catalog.Value.Entries).Add("Mutate", catalog.Value));
    }

    [Fact]
    public void RawStructure_EnforcesProjectionBoundsAndTrySurface() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text(new string('x', 200)))
            .ToBytes();
        PdfDocument document = PdfDocument.Open(bytes);

        PdfRawDocumentView raw = document.Read.RawStructure(new PdfRawStructureOptions {
            MaxObjects = 2,
            MaxDepth = 1,
            MaxCollectionItems = 1,
            MaxTextLength = 8
        });
        PdfOperationResult<PdfRawDocumentView> attempted = document.Read.TryRawStructure();

        Assert.Equal(2, raw.Objects.Count);
        Assert.True(raw.IsTruncated);
        Assert.True(raw.TrailerPreview.Length <= 8);
        Assert.True(attempted.Succeeded, string.Join(Environment.NewLine, attempted.Diagnostics));
    }
}
