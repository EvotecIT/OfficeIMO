using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTaggedPageOrganizationTests {
    [Fact]
    public void DuplicateRetargetsDirectAnnotationPageOwners() {
        byte[] source = CreateTaggedSource();
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var parsed = PdfReadDocument.Open(source);
            int pageNumber = parsed.Pages[0].ObjectNumber;
            var page = (PdfDictionary)objects[pageNumber].Value;
            var note = new PdfDictionary(); note.Items["Type"] = new PdfName("Annot"); note.Items["Subtype"] = new PdfName("Text");
            note.Items["P"] = new PdfReference(pageNumber, 0); note.Items["Contents"] = new PdfStringObj("Direct note");
            var rect = new PdfArray(); foreach (double value in new[] { 10D, 10D, 30D, 30D }) rect.Items.Add(new PdfNumber(value));
            note.Items["Rect"] = rect;
            var annotations = new PdfArray(); annotations.Items.Add(note); page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });
        byte[] output = PdfDocument.Load(source).Pages.Duplicate(1).ToBytes();
        var read = PdfReadDocument.Open(output);
        foreach (var pageInfo in read.Pages.Take(2)) {
            var page = (PdfDictionary)read.Objects[pageInfo.ObjectNumber].Value;
            var annotations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(read.Objects, page.Items["Annots"]));
            var note = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(read.Objects, Assert.Single(annotations.Items)));
            Assert.Equal(pageInfo.ObjectNumber, Assert.IsType<PdfReference>(note.Items["P"]).ObjectNumber);
        }
    }
}
