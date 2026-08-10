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

        HtmlRoundTripScore typedValueScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"date\" value=\"not-a-date\"><input type=\"color\" value=\"red\"></form></main>",
            "<main><form><input type=\"date\" value=\"\"><input type=\"color\" value=\"#000000\"></form></main>");
        Assert.Equal(1D, typedValueScore.Metrics["form-state"], 3);

        HtmlRoundTripScore defaultColorScore = HtmlRoundTripScorer.Compare(
            "<main><input type=\"color\"></main>",
            "<main><input type=\"color\" value=\"#000000\"></main>");
        Assert.Equal(1D, defaultColorScore.Metrics["form-state"], 3);

        HtmlRoundTripScore effectiveRadioScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"radio\" name=\"choice\" checked><input type=\"radio\" name=\"choice\" checked></form></main>",
            "<main><form><input type=\"radio\" name=\"choice\"><input type=\"radio\" name=\"choice\" checked></form></main>");
        Assert.Equal(1D, effectiveRadioScore.Metrics["form-state"], 3);

        HtmlRoundTripScore steppedRangeScore = HtmlRoundTripScorer.Compare(
            "<main><input type=\"range\" min=\"0\" max=\"10\" step=\"3\" value=\"8\"></main>",
            "<main><input type=\"range\" min=\"0\" max=\"10\" step=\"3\" value=\"9\"></main>");
        Assert.Equal(1D, steppedRangeScore.Metrics["form-state"], 3);
    }

    [Fact]
    public void FidelityScore_IgnoresInapplicableControlAttributesButRetainsOptionSelectedness() {
        HtmlRoundTripScore inapplicableAttributeScore = HtmlRoundTripScorer.Compare(
            """
            <main><input selected><select type="menu" value="forged"><option type="fake" name="forged">First</option><option selected>Second</option></select><textarea type="text" value="forged">Body</textarea><button autocomplete="off">Go</button></main>
            """,
            """
            <main><input><select><option>First</option><option selected>Second</option></select><textarea>Body</textarea><button>Go</button></main>
            """);
        Assert.Equal(1D, inapplicableAttributeScore.Metrics["form-state"], 3);

        HtmlRoundTripScore effectiveSelectionScore = HtmlRoundTripScorer.Compare(
            "<main><select><option>First</option><option selected>Second</option></select></main>",
            "<main><select><option selected>First</option><option>Second</option></select></main>");
        Assert.True(effectiveSelectionScore.Metrics["form-state"] < 1D);
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
