using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrOrientationTests {
    [Theory]
    [InlineData(0, "15.0", 1D)]
    [InlineData(90, "7.5", 0.5D)]
    [InlineData(180, "28.28", 1D)]
    [InlineData(270, "9.30", 0.62D)]
    public void OrientationParserReadsCorrectiveRotationAndCalibratedEvidence(int rotation, string confidence, double expected) {
        OcrOrientationResult? result = TesseractOcrEngine.ParseOrientation(
            "Page number: 0\r\nOrientation in degrees: 90\r\nRotate: " + rotation +
            "\r\nOrientation confidence: " + confidence + "\r\nScript: Latin\r\nScript confidence: 8.06\r\n");
        Assert.NotNull(result);
        Assert.Equal(rotation, result!.ClockwiseRotationDegrees);
        Assert.Equal(expected, result.Confidence, 6);
        Assert.Equal("Latin", result.Script);
    }

    [Theory]
    [InlineData("Rotate: 45\nOrientation confidence: 15")]
    [InlineData("Rotate: 90\nOrientation confidence: NaN")]
    [InlineData("Rotate: 90\nOrientation confidence: -1")]
    [InlineData("Rotate: 90\nOrientation confidence: Infinity")]
    [InlineData("Rotate: 90")]
    [InlineData("Rotate: 90\nRotate: 270\nOrientation confidence: 15")]
    [InlineData("Rotate: 90\nOrientation confidence: 15\nOrientation confidence: invalid")]
    public void OrientationParserRejectsIncompleteInvalidOrConflictingEvidence(string output) {
        Assert.Null(TesseractOcrEngine.ParseOrientation(output));
    }
}
