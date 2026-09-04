using OfficeIMO.Ocr.Tesseract;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TesseractOcrFacadeTests {
    [Theory]
    [InlineData(TesseractOcrLanguage.English, "eng")]
    [InlineData(TesseractOcrLanguage.Polish, "pol")]
    [InlineData(TesseractOcrLanguage.English | TesseractOcrLanguage.Polish, "eng+pol")]
    [InlineData(TesseractOcrLanguage.French | TesseractOcrLanguage.German | TesseractOcrLanguage.Spanish, "fra+deu+spa")]
    [InlineData(TesseractOcrLanguage.Arabic | TesseractOcrLanguage.Hebrew | TesseractOcrLanguage.Hindi, "ara+heb+hin")]
    [InlineData(TesseractOcrLanguage.ChineseSimplified | TesseractOcrLanguage.Japanese | TesseractOcrLanguage.Korean, "chi_sim+jpn+kor")]
    public void Languages_MapDiscoverableValuesToStableProviderExpressions(TesseractOcrLanguage languages, string expected) {
        Assert.Equal(expected, languages.ToTesseractExpression());
    }

    [Fact]
    public void Languages_RejectEmptyAndUndefinedSelections() {
        Assert.Throws<ArgumentOutOfRangeException>(() => ((TesseractOcrLanguage)0).ToTesseractExpression());
        Assert.Throws<ArgumentOutOfRangeException>(() => ((TesseractOcrLanguage)(1UL << 40)).ToTesseractExpression());
    }

    [Fact]
    public void Languages_ExposeOneTypedEntryForEveryProvisionedModel() {
        Assert.Equal(28, TesseractOcrLanguages.Supported.Count);
        Assert.Equal(TesseractOcrLanguages.Supported.Count, TesseractOcrLanguages.Supported.Distinct().Count());
        Assert.All(TesseractOcrLanguages.Supported, language => {
            string code = language.ToTesseractExpression();
            Assert.Contains(code, TesseractLanguageData.SupportedLanguages);
        });
    }

    [Fact]
    public void Languages_PreserveEngineConfigurationAndRejectAmbiguousOverrides() {
        Assert.Equal("eng", TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions()));
        Assert.Equal("eng+pol", TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions {
            Languages = TesseractOcrLanguage.English | TesseractOcrLanguage.Polish
        }));
        Assert.Equal("pol", TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions {
            Engine = new TesseractOcrEngineOptions { Language = "pol" }
        }));
        Assert.Equal("deu", TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions {
            CustomLanguageExpression = "deu"
        }));

        Assert.Throws<ArgumentException>(() => TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions {
            Languages = TesseractOcrLanguage.English | TesseractOcrLanguage.Polish,
            Engine = new TesseractOcrEngineOptions { Language = "pol" }
        }));
        Assert.Throws<ArgumentException>(() => TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions {
            CustomLanguageExpression = "deu",
            Engine = new TesseractOcrEngineOptions { Language = "pol" }
        }));
        Assert.Throws<ArgumentException>(() => TesseractOcr.ResolveLanguageExpression(new TesseractOcrSessionOptions {
            Languages = TesseractOcrLanguage.Polish,
            CustomLanguageExpression = "deu"
        }));
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(11, false)]
    [InlineData(12, true)]
    [InlineData(13, false)]
    public void Session_RequiresOrientationDataOnlyForOsdSegmentationModes(int? pageSegmentationMode, bool expectsOsd) {
        string[] required = TesseractOcr.ResolveRequiredLanguageData("eng+pol", pageSegmentationMode);

        Assert.Equal(expectsOsd, required.Contains("osd", StringComparer.Ordinal));
        Assert.Contains("eng", required);
        Assert.Contains("pol", required);
        Assert.Equal(required.Length, required.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public async Task RecognizeFileAsync_RejectsNullNestedEngineOptionsBeforeReadingTheFile() {
        var options = new TesseractOcrSessionOptions { Engine = null! };

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            TesseractOcr.RecognizeFileAsync("missing-image.png", options));

        Assert.Equal("options", exception.ParamName);
        Assert.Contains("engine options", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
