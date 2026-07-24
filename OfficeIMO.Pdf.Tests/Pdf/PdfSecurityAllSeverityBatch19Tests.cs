using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSecurityAllSeverityBatch19Tests {
    [Fact]
    public void AnnotationUpdateRejectsNullInkPathEntryAtPublicBoundary() {
        byte[] pdf = PdfDocument.Create()
            .TextAnnotation("note")
            .Paragraph(paragraph => paragraph.Text("source"))
            .ToBytes();
        int objectNumber = Assert.Single(
            PdfInspector.Inspect(pdf).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
        var paths = new IReadOnlyList<double>[] { null! };

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            PdfAnnotationEditor.UpdateAnnotation(
                pdf,
                objectNumber,
                new PdfAnnotationUpdateOptions { InkPaths = paths }));

        Assert.Contains("cannot contain null", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BookmarkEditorRejectsNonFiniteDestinationCoordinates() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("one page"))
            .ToBytes();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfBookmarkEditor.Edit(
                source,
                session => session.Add("invalid", 1, destinationTop: double.NaN)));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfBookmarkEditor.Edit(source, session => {
                PdfBookmarkNode node = session.Add("valid", 1);
                session.Retarget(
                    node.Id,
                    1,
                    PdfOpenActionDestinationMode.Xyz,
                    destinationZoom: double.PositiveInfinity);
            }));
    }
}
