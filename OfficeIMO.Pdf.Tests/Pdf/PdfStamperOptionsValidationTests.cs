using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStamperTests {
    [Fact]
    public void TextStampOptions_SnapshotPageNumbersAndRejectInvalidValues() {
        var pageNumbers = new[] { 1, 2 };
        var options = new PdfTextStampOptions {
            PageNumbers = pageNumbers
        };

        pageNumbers[0] = 99;
        int[] readback = options.PageNumbers!;
        readback[1] = 88;

        Assert.Equal(new[] { 1, 2 }, options.PageNumbers);

        var ranged = new PdfTextStampOptions();
        Assert.Same(ranged, ranged.UsePageRange(2, 2));
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);
        int[] rangedReadback = ranged.PageNumbers!;
        rangedReadback[0] = 99;
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);

        var modelRange = new PdfTextStampOptions();
        Assert.Same(modelRange, modelRange.UsePageRange(PdfPageRange.From(1, 2)));
        Assert.Equal(new[] { 1, 2 }, modelRange.PageNumbers);

        var rangeList = new PdfTextStampOptions();
        Assert.Same(rangeList, rangeList.UsePageRanges(PdfPageRange.ParseMany("1-2,2")));
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);
        int[] rangeListReadback = rangeList.PageNumbers!;
        rangeListReadback[0] = 99;
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { PageNumbers = new[] { 0 } });
        Assert.Throws<ArgumentException>(() => new PdfTextStampOptions { PageNumbers = new[] { 1, 1 } });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions().UsePageRange(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions().UsePageRange(2, 1));
        Assert.Throws<ArgumentNullException>(() => new PdfTextStampOptions().UsePageRanges(null!));
        Assert.Throws<ArgumentException>(() => new PdfTextStampOptions().UsePageRanges(Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { X = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { Y = double.PositiveInfinity });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { FontSize = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { FontSize = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { RotationDegrees = double.NegativeInfinity });
        var fontException = Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { Font = (PdfStandardFont)99 });
        Assert.Equal("Font", fontException.ParamName);
        Assert.Contains("Text stamp font must be one of the supported standard PDF fonts.", fontException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageStampOptions_SnapshotPageNumbersAndRejectInvalidValues() {
        var pageNumbers = new[] { 1, 2 };
        var options = new PdfImageStampOptions {
            PageNumbers = pageNumbers
        };

        pageNumbers[0] = 99;
        int[] readback = options.PageNumbers!;
        readback[1] = 88;

        Assert.Equal(new[] { 1, 2 }, options.PageNumbers);

        var ranged = new PdfImageStampOptions();
        Assert.Same(ranged, ranged.UsePageRange(2, 2));
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);
        int[] rangedReadback = ranged.PageNumbers!;
        rangedReadback[0] = 99;
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);

        var modelRange = new PdfImageStampOptions();
        Assert.Same(modelRange, modelRange.UsePageRange(PdfPageRange.From(1, 2)));
        Assert.Equal(new[] { 1, 2 }, modelRange.PageNumbers);

        var rangeList = new PdfImageStampOptions();
        Assert.Same(rangeList, rangeList.UsePageRanges(PdfPageRange.ParseMany("1-2,2")));
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);
        int[] rangeListReadback = rangeList.PageNumbers!;
        rangeListReadback[0] = 99;
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { PageNumbers = new[] { 0 } });
        Assert.Throws<ArgumentException>(() => new PdfImageStampOptions { PageNumbers = new[] { 1, 1 } });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions().UsePageRange(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions().UsePageRange(2, 1));
        Assert.Throws<ArgumentNullException>(() => new PdfImageStampOptions().UsePageRanges(null!));
        Assert.Throws<ArgumentException>(() => new PdfImageStampOptions().UsePageRanges(Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { X = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Y = double.PositiveInfinity });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Width = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Width = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Height = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Height = double.PositiveInfinity });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { RotationDegrees = double.NegativeInfinity });
    }

    [Fact]
    public void Stamper_RejectsInvalidInputs() {
        byte[] source = BuildTwoPagePdf();

        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText((byte[])null!, "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText((Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(new WriteOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText(source, null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(source, string.Empty));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { FontSize = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { PageNumbers = new[] { 0 } }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { PageNumbers = new[] { 3 } }));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { PageNumbers = new[] { 1, 1 } }));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText(source, null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(source, new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText(new MemoryStream(source), (string)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(new MemoryStream(source), " ", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampTextToBytes(" ", "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText("input.pdf", (Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText("missing.pdf", new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(" ", new MemoryStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(" ", "out.pdf", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText("missing.pdf", " ", "x"));

        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText((byte[])null!, "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText((Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(new WriteOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText(source, null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(source, string.Empty));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText(source, null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(source, new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText(new MemoryStream(source), (string)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(new MemoryStream(source), " ", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkTextToBytes(" ", "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText("input.pdf", (Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText("missing.pdf", new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(" ", new MemoryStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(" ", "out.pdf", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText("missing.pdf", " ", "x"));

        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage((byte[])null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage((Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new WriteOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(source, (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(new MemoryStream(source), (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new MemoryStream(source), new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(source, (byte[])null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, Array.Empty<byte>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { Width = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { Height = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { PageNumbers = new[] { 0 } }));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { PageNumbers = new[] { 1, 1 } }));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(source, null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(new MemoryStream(source), (string)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new MemoryStream(source), " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new MemoryStream(source), " ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImageToBytes(" ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImageToBytes(" ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage("input.pdf", (Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", new MemoryStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage("input.pdf", (Stream)null!, new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", new ReadOnlyStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", new MemoryStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", "out.pdf", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", "out.pdf", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", " ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage((byte[])null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage((Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new WriteOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(source, (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(source, new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(source, (byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(source, null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(source, new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), (string)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), " ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImageToBytes(" ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImageToBytes(" ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage("input.pdf", (Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", new MemoryStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage("input.pdf", (Stream)null!, new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", new ReadOnlyStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", new MemoryStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", "out.pdf", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", "out.pdf", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", " ", new MemoryStream(CreateMinimalRgbPng())));
    }

    [Fact]
    public void Stamper_PathOutputsRejectDirectoryTargets() {
        byte[] source = BuildTwoPagePdf();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-output-path-" + Guid.NewGuid().ToString("N"));

        try {
            Directory.CreateDirectory(directory);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfStamper.WatermarkText(new MemoryStream(source), directory, "DRAFT"));

            Assert.Equal("outputPath", exception.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", exception.Message, StringComparison.Ordinal);

            var pathInputException = Assert.Throws<ArgumentException>(() =>
                PdfStamper.StampText("missing.pdf", directory, "DRAFT"));

            Assert.Equal("outputPath", pathInputException.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", pathInputException.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
