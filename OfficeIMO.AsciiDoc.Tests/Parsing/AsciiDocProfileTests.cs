namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocProfileTests {
    [Theory]
    [InlineData(AsciiDocDocumentProfile.OfficeIMO)]
    [InlineData(AsciiDocDocumentProfile.PreserveOnly)]
    public void NamedParseProfilesAreExplicitAndLossless(AsciiDocDocumentProfile profile) {
        const string source = "= Profile\n\nNOTE: Typed content\ninclude::part.adoc[]\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(
            source,
            AsciiDocParseOptions.CreateProfile(profile)).Document;

        Assert.Equal(profile, document.Profile);
        Assert.Equal(source, document.ToAsciiDoc());
        Assert.Contains(document.Blocks, block => block is AsciiDocAdmonitionBlock);
        Assert.Contains(document.Blocks, block => block is AsciiDocBlockMacro macro && macro.Name == "include");
    }

    [Fact]
    public void NamedProcessingProfileDoesNotEnableIncludesOrExtensions() {
        const string source = "include::part.adoc[]\ncustom::value[]\n";
        AsciiDocProcessorOptions options = AsciiDocProcessorOptions.CreateProfile(AsciiDocDocumentProfile.OfficeIMO);

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(source, options);

        Assert.Equal(AsciiDocDocumentProfile.OfficeIMO, result.SourceDocument.Profile);
        Assert.Equal(AsciiDocDocumentProfile.OfficeIMO, result.Document.Profile);
        Assert.Equal(source, result.ProcessedSource);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCPROC001");
        Assert.Null(options.IncludeResolver);
        Assert.Null(options.Extensions);
    }

    [Fact]
    public void UnknownProfilesAreRejected() {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            AsciiDocParseOptions.CreateProfile((AsciiDocDocumentProfile)999));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            AsciiDocProcessorOptions.CreateProfile((AsciiDocDocumentProfile)999));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            AsciiDocDocument.Parse("= Invalid profile", new AsciiDocParseOptions {
                Profile = (AsciiDocDocumentProfile)999
            }));
    }
}
