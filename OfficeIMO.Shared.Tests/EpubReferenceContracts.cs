using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubReferenceContractTests {
    [Fact]
    public void Resolve_PreservesQueryAndFragmentWhileResolvingEncodedContainerPath() {
        EpubReference reference = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "../images/cover%20art.png?size=large#hero%20image");

        Assert.True(reference.IsValid);
        Assert.Equal(EpubReferenceKind.Container, reference.Kind);
        Assert.Equal(EpubReferenceError.None, reference.Error);
        Assert.Equal("EPUB/images/cover art.png", reference.ContainerPath);
        Assert.Equal("EPUB/images/cover%20art.png", reference.ContainerUrlPath);
        Assert.Equal("size=large", reference.Query);
        Assert.Equal("hero image", reference.Fragment);
        Assert.Equal("EPUB/images/cover%20art.png?size=large", reference.Target);
        Assert.Equal("EPUB/images/cover%20art.png?size=large#hero%20image", reference.ResolvedValue);
        Assert.True(reference.IsConforming);
    }

    [Theory]
    [InlineData("../images/cover%23v2.png", "EPUB/images/cover#v2.png", "EPUB/images/cover%23v2.png")]
    [InlineData("../images/cover%3Fv2.png", "EPUB/images/cover?v2.png", "EPUB/images/cover%3Fv2.png")]
    public void Resolve_SeparatesArchiveLookupPathFromUrlSerialization(
        string value,
        string expectedContainerPath,
        string expectedResolvedValue) {
        EpubReference reference = EpubReference.Resolve("EPUB/text/chapter.xhtml", value);

        Assert.Equal(EpubReferenceKind.Container, reference.Kind);
        Assert.Equal(expectedContainerPath, reference.ContainerPath);
        Assert.Equal(expectedResolvedValue, reference.ContainerUrlPath);
        Assert.Equal(expectedResolvedValue, reference.ResolvedValue);
    }

    [Theory]
    [InlineData("#section", "EPUB/text/chapter.xhtml#section")]
    [InlineData("?mode=print", "EPUB/text/chapter.xhtml?mode=print")]
    [InlineData("?mode=print#section", "EPUB/text/chapter.xhtml?mode=print#section")]
    public void Resolve_BindsFragmentAndQueryOnlyReferencesToCurrentDocument(string value, string expected) {
        EpubReference reference = EpubReference.Resolve("EPUB/text/chapter.xhtml", value);

        Assert.Equal(EpubReferenceKind.Container, reference.Kind);
        Assert.Equal("EPUB/text/chapter.xhtml", reference.ContainerPath);
        Assert.Equal(expected, reference.ResolvedValue);
    }

    [Fact]
    public void Resolve_ToleratesButMarksContainerRootRelativeReferenceNonConforming() {
        EpubReference reference = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "/EPUB/images/cover.png#front");

        Assert.Equal(EpubReferenceKind.Container, reference.Kind);
        Assert.Equal("EPUB/images/cover.png", reference.ContainerPath);
        Assert.True(reference.IsContainerRootRelative);
        Assert.False(reference.IsConforming);
    }

    [Theory]
    [InlineData("../../../outside.xhtml", EpubReferenceError.EscapesContainer)]
    [InlineData("../images/a%2Fb.png", EpubReferenceError.InvalidPath)]
    [InlineData("../images/a%5Cb.png", EpubReferenceError.InvalidPath)]
    [InlineData("../images/bad%2.png", EpubReferenceError.InvalidPath)]
    [InlineData("file:///tmp/outside.xhtml", EpubReferenceError.FileUrl)]
    public void Resolve_RejectsUnsafeOrAmbiguousContainerReferences(string value, EpubReferenceError error) {
        EpubReference reference = EpubReference.Resolve("EPUB/text/chapter.xhtml", value);

        Assert.False(reference.IsValid);
        Assert.Equal(EpubReferenceKind.Invalid, reference.Kind);
        Assert.Equal(error, reference.Error);
        Assert.Null(reference.ResolvedValue);
    }

    [Fact]
    public void Resolve_AppliesUrlDotSegmentsBeforeDecodingArchiveNames() {
        EpubReference reference = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "%2e%2e/images/cover.png");

        Assert.Equal(EpubReferenceKind.Container, reference.Kind);
        Assert.Equal("EPUB/images/cover.png", reference.ContainerPath);
    }

    [Theory]
    [InlineData("https://cdn.example/book/audio.mp3?token=1#clip", EpubReferenceKind.External)]
    [InlineData("//cdn.example/book/audio.mp3", EpubReferenceKind.External)]
    [InlineData("data:image/svg+xml,%3Csvg%2F%3E", EpubReferenceKind.Data)]
    public void Resolve_ClassifiesNonContainerReferencesWithoutFetching(string value, EpubReferenceKind kind) {
        EpubReference reference = EpubReference.Resolve("EPUB/text/chapter.xhtml", value);

        Assert.Equal(kind, reference.Kind);
        Assert.Equal(value, reference.ResolvedValue);
        Assert.Null(reference.ContainerPath);
    }

    [Fact]
    public void Resolve_AppliesRelativeAndExternalHtmlBaseHref() {
        EpubReference local = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "../shared/",
            "images/cover.png#front");
        EpubReference external = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "https://cdn.example/book/",
            "audio/chapter.mp3");

        Assert.Equal("EPUB/shared/images/cover.png#front", local.ResolvedValue);
        Assert.Equal(EpubReferenceKind.External, external.Kind);
        Assert.Equal("https://cdn.example/book/audio/chapter.mp3", external.ResolvedValue);
    }

    [Theory]
    [InlineData("/", false)]
    [InlineData("./", true)]
    public void Resolve_AppliesContainerRootHtmlBaseHref(string baseHref, bool isConforming) {
        EpubReference reference = EpubReference.Resolve(
            "chapter.xhtml",
            baseHref,
            "images/cover.png");

        Assert.Equal(EpubReferenceKind.Container, reference.Kind);
        Assert.Equal(EpubReferenceError.None, reference.Error);
        Assert.Equal("images/cover.png", reference.ContainerPath);
        Assert.Equal("images/cover.png", reference.ResolvedValue);
        Assert.Equal(isConforming, reference.IsConforming);
    }

    [Fact]
    public void Resolve_BindsFragmentOnlyReferenceToDirectoryBaseAndPreservesBaseQuery() {
        EpubReference directory = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "../shared/?edition=2",
            "#notes");
        EpubReference document = EpubReference.Resolve(
            "EPUB/text/chapter.xhtml",
            "appendix.xhtml?edition=2",
            "?edition=3#notes");

        Assert.Equal("EPUB/shared/?edition=2#notes", directory.ResolvedValue);
        Assert.Equal("EPUB/text/appendix.xhtml?edition=3#notes", document.ResolvedValue);
    }
}
