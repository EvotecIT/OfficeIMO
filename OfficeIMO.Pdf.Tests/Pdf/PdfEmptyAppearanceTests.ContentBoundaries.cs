using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfEmptyAppearanceTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FlatteningPreservesPagePaintAndAppearanceState(bool form, bool leaveChangedGraphicsState) {
        string pageContent = "0 g 10 160 50 20 re f" + (leaveChangedGraphicsState ? "\n1 0 0 1 40 0 cm\n" : "");
        byte[] source = BuildPdf(form ? "[0 0 20 20]" : null, "[0 0 20 20]", pageContent, widgetOn: form);
        byte[] flattened = form
            ? PdfDocument.Load(source).Forms.Flatten().ToBytes()
            : PdfDocument.Load(source).Annotations.Flatten().Bytes;

        Assert.Equal(Render(source), Render(flattened));
        AssertOriginalPagePaintBytes(source, flattened);
    }

    [Fact]
    public void TextReplacementStampPreservesPagePaintWhenTheOriginalStreamHasNoTrailingWhitespace() {
        byte[] source = BuildPdf(null, "[0 0 0 0]");
        var request = new PdfStamper.TextStampRequest(1, "boundary audit", 20D, 20D,
            PdfStandardFont.Helvetica, 10D, new PdfColor(0D, 0D, 0D), 0D, textRenderingMode: 3);

        byte[] stamped = PdfStamper.StampTextBatch(source, new[] { request });

        Assert.Equal(Render(source), Render(stamped));
        AssertOriginalPagePaintBytes(source, stamped);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(stamped);
        Assert.Contains("boundary audit", independent.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ImportedPagePreservesTheAdmittedSplitNumericOperandAndItsPaint() {
        byte[] unsplit = BuildPdf(null, "[0 0 0 0]");
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(unsplit, null, null, (objects, security) => {
            var first = (PdfStream)objects[8].Value;
            objects[8] = new PdfIndirectObject(8, 0, new PdfStream(first.Dictionary, Encoding.ASCII.GetBytes("0 g 10 1")));
            int next = objects.Keys.Max() + 1;
            objects[next] = new PdfIndirectObject(next, 0, new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("60 50 20 re f")));
            var contents = new PdfArray();
            contents.Items.Add(new PdfReference(8, 0));
            contents.Items.Add(new PdfReference(next, 0));
            ((PdfDictionary)objects[3].Value).Items["Contents"] = contents;
            return security.InfoObjectNumber;
        });
        PdfDocument input = PdfDocument.Load(source);
        PdfDocument target = PdfDocument.Load(BuildPdf(null, "[0 0 0 0]", pageContent: ""));

        PdfDocument imported = target.Stamp.OverlayPage(source, new PdfPageOverlayOptions {
            X = 0D, Y = 0D, Fit = PdfPageOverlayFit.None
        });

        Assert.Equal(Render(unsplit), Render(source));
        Assert.Equal(Render(source), Render(imported.ToBytes()));
        Assert.Equal(source, input.ToBytes());
    }

    [Fact]
    public void FlatteningAdmitsOnlyTheKnownAddedContentBeyondTightSourceLimits() {
        byte[] source = BuildPdf(null, "[0 0 20 20]");
        PdfStream[] streams = PdfSyntax.ParseObjects(source).Map.Values.Select(item => item.Value)
            .OfType<PdfStream>().ToArray();
        int sourcePageBytes = ((PdfStream)PdfSyntax.ParseObjects(source).Map[8].Value).Data.Length;
        var limits = new PdfReadLimits {
            MaxInputBytes = source.LongLength,
            MaxRawStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxDecodedStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxTotalDecodedStreamBytes = streams.Sum(stream => (long)stream.Data.Length),
            MaxPageContentBytes = sourcePageBytes,
            MaxRetainedContentBytes = streams.Sum(stream => (long)stream.Data.Length),
            MaxContentOperations = 3,
            MaxContentOperands = 7
        };

        PdfAnnotationEditResult result = PdfDocument.Load(source, new PdfLoadOptions { Limits = limits }).Annotations.Flatten();
        PdfDocument flattened = result.ToDocument();

        Assert.Equal(Render(source), Render(flattened.ToBytes()));
        Assert.Empty(flattened.Reader.Annotations());
        Assert.True(flattened.ReadOptions.Limits.MaxPageContentBytes < new PdfReadLimits().MaxPageContentBytes);
        Assert.True(flattened.ReadOptions.Limits.MaxContentOperations < new PdfReadLimits().MaxContentOperations);
    }

    private static byte[] Render(byte[] pdf) {
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ContinueOnError = false }));
        Assert.True(result.Succeeded);
        return Assert.IsType<byte[]>(result.Bytes);
    }

    private static void AssertOriginalPagePaintBytes(byte[] source, byte[] output) {
        PdfStream original = Assert.IsType<PdfStream>(PdfSyntax.ParseObjects(source).Map[8].Value);
        Assert.Contains(PdfSyntax.ParseObjects(output).Map.Values, item => item.Value is PdfStream stream
            && stream.Data.SequenceEqual(original.Data));
    }
}
