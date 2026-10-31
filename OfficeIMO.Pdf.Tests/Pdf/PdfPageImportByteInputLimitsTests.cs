using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageImportByteInputLimitsTests {
    [Theory]
    [InlineData("append", false)]
    [InlineData("prepend", false)]
    [InlineData("insert", false)]
    [InlineData("append", true)]
    [InlineData("prepend", true)]
    [InlineData("insert", true)]
    public void AllPagesByteImportsEnforceTheOriginalInputLimit(string placement, bool flatten) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Incoming")).ToBytes();
        byte[] targetBytes = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Original")).ToBytes();
        var target = PdfDocument.Load(targetBytes);
        var options = new PdfPageImportOptions {
            FlattenVisualAnnotations = flatten,
            SourceReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length - 1 } }
        };

        var error = Assert.Throws<PdfReadLimitException>(() => placement switch {
            "append" => target.Pages.Append(source, options),
            "prepend" => target.Pages.Prepend(source, options),
            _ => target.Pages.Insert(1, source, options)
        });

        Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
        Assert.Equal(source.Length - 1, error.Limit);
        Assert.Equal(source.Length, error.Actual);
        Assert.Equal(targetBytes, target.ToBytes());
    }
}
