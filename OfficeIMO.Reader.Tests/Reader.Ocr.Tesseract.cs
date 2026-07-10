using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Tesseract;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrTesseractTests {
    [Fact]
    public void TesseractTsvParser_MapsLinesWordsConfidenceAndPixelGeometry() {
        const string tsv = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n"
            + "1\t1\t0\t0\t0\t0\t0\t0\t200\t100\t-1\t\n"
            + "5\t1\t1\t1\t1\t1\t10\t20\t40\t10\t90\tInvoice\n"
            + "5\t1\t1\t1\t1\t2\t55\t20\t30\t10\t80\t1042\n"
            + "5\t1\t1\t1\t2\t1\t10\t40\t50\t12\t100\tTotal\n";

        OfficeOcrEngineResult result = TesseractTsvParser.Parse(tsv, "eng");

        Assert.Equal("Invoice 1042" + Environment.NewLine + "Total", result.Text);
        Assert.Equal(0.9D, result.Confidence!.Value, precision: 6);
        Assert.Equal(2, result.Spans.Count(span => span.Level == OfficeOcrTextSpanLevel.Line));
        Assert.Equal(3, result.Spans.Count(span => span.Level == OfficeOcrTextSpanLevel.Word));
        OfficeOcrTextSpan firstLine = Assert.Single(result.Spans, span => span.Level == OfficeOcrTextSpanLevel.Line && span.Text == "Invoice 1042");
        Assert.Equal(10D, firstLine.Region!.X);
        Assert.Equal(75D, firstLine.Region.Width);
        Assert.Equal(OfficeOcrCoordinateUnit.Pixels, firstLine.CoordinateUnit);
        Assert.Equal("tesseract-cli", result.Provider);
    }

    [Fact]
    public void TesseractOcrEngine_BuildsOptionsBeforeTsvOutputConfig() {
        var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
            TessdataDirectory = "/models",
            EngineMode = 1,
            PageSegmentationMode = 6,
            Dpi = 300,
            AdditionalArguments = new[] { "quiet" }
        });

        IReadOnlyList<string> arguments = engine.BuildRecognitionArguments("input image.png", "result", "eng+pol");

        Assert.Equal("input image.png", arguments[0]);
        Assert.Equal("result", arguments[1]);
        Assert.Contains("eng+pol", arguments);
        Assert.Contains("/models", arguments);
        Assert.Contains("300", arguments);
        Assert.Equal("quiet", arguments[arguments.Count - 2]);
        Assert.Equal("tsv", arguments[arguments.Count - 1]);
    }
}
