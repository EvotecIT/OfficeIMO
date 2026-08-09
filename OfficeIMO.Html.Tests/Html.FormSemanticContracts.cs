using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void FidelityScore_NormalizesRangeCurrentValuesAndInvalidFormKeywords() {
        HtmlRoundTripScore rangeDefaultValueScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"range\"></form></main>",
            "<main><form><input type=\"range\" value=\"50\"></form></main>");
        Assert.Equal(1D, rangeDefaultValueScore.Metrics["form-state"], 3);

        HtmlRoundTripScore rangeSanitizedValueScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"range\" min=\"10\" max=\"20\" value=\"200\"></form></main>",
            "<main><form><input type=\"range\" min=\"10\" max=\"20\" value=\"20\"></form></main>");
        Assert.Equal(1D, rangeSanitizedValueScore.Metrics["form-state"], 3);

        HtmlRoundTripScore whitespacePaddedKeywordScore = HtmlRoundTripScorer.Compare(
            "<main><form method=\" post \" enctype=\" multipart/form-data \"><input type=\" checkbox \" checked></form></main>",
            "<main><form method=\"get\" enctype=\"application/x-www-form-urlencoded\"><input type=\"text\"></form></main>");
        Assert.Equal(1D, whitespacePaddedKeywordScore.Metrics["form-state"], 3);
    }

    [Fact]
    public void ResourceManifest_UsesEffectiveFormControlTypes() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <input type=" image " src="file:///secret/image.png" data-src="file:///secret/lazy.png" formaction="file:///secret/input-action">
            <button type=" reset " formaction="file:///secret/button-action">Go</button>
            """);

        Assert.DoesNotContain(manifest.Resources, resource =>
            resource.ElementName == "input" &&
            (resource.AttributeName == "src" || resource.AttributeName == "data-src" || resource.AttributeName == "formaction"));
        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "button" &&
            resource.AttributeName == "formaction" &&
            resource.Source == "file:///secret/button-action");
    }
}
