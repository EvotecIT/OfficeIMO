using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FormPathOutputStreams_RejectNullAndReadOnlyOutputsBeforeReadingInputs() {
        var values = new Dictionary<string, string> {
            ["Name"] = "Value"
        };

        Assert.Throws<ArgumentNullException>(() => PdfFormFiller.FillFields("input.pdf", (Stream)null!, values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields("missing.pdf", new ReadOnlyStream(), values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(" ", new MemoryStream(), values));
        Assert.Throws<ArgumentNullException>(() => PdfFormFiller.FlattenFields("input.pdf", (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FlattenFields("missing.pdf", new ReadOnlyStream()));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FlattenFields(" ", new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfFormFiller.FillAndFlattenFields("input.pdf", (Stream)null!, values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillAndFlattenFields("missing.pdf", new ReadOnlyStream(), values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillAndFlattenFields(" ", new MemoryStream(), values));
    }

    [Fact]
    public void FillFields_RejectsUnknownFieldNames() {
        var ex = Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(BuildHierarchicalFormPdf(), new Dictionary<string, string> {
            ["Missing"] = "Value"
        }));

        Assert.Contains("PDF form field was not found: Missing", ex.Message);
    }

    [Fact]
    public void FillFields_RejectsSignedPdfs() {
        var ex = Assert.Throws<NotSupportedException>(() => PdfFormFiller.FillFields(BuildSignedFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Value"
        }));

        Assert.Equal("Signed PDF files are not supported for form filling by OfficeIMO.Pdf yet.", ex.Message);
    }

    [Fact]
    public void FlattenFields_RejectsSignedPdfs() {
        var ex = Assert.Throws<NotSupportedException>(() => PdfFormFiller.FlattenFields(BuildSignedFormPdf()));

        Assert.Equal("Signed PDF files are not supported for form flattening by OfficeIMO.Pdf yet.", ex.Message);
    }
}
