using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionTextObjectScopeTests {
    [Fact]
    public void MatchesExpectedSurvivorsAcrossClonedContentStreamOwners() {
        PdfRedactionTextObjectScope reviewed = CreateReviewedScope(owner: 10);
        PdfRedactionTextObjectScope current = CreateSurvivorScope(owner: 20);

        Assert.True(reviewed.Matches(current));
    }

    [Fact]
    public void SurvivorIdentityIncludesEncodedGlyphBytes() {
        PdfRedactionTextObjectScope reviewed = CreateReviewedScope(owner: 10);
        PdfRedactionTextObjectScope current = CreateSurvivorScope(owner: 20, glyphByte: 0x43);

        Assert.False(reviewed.Matches(current));
    }

    [Fact]
    public void SurvivorIdentityIncludesClipPath() {
        PdfRedactionTextObjectScope reviewed = CreateReviewedScope(owner: 10);
        PdfRedactionTextObjectScope current = CreateSurvivorScope(
            owner: 20,
            clipPath: PdfPageClipPath.Rectangle(0D, 0D, 40D, 40D));

        Assert.False(reviewed.Matches(current));
    }

    [Fact]
    public void SurvivorIdentityIncludesFullTextTransform() {
        PdfRedactionTextObjectScope reviewed = CreateReviewedScope(owner: 10);
        PdfRedactionTextObjectScope current = CreateSurvivorScope(
            owner: 20,
            transform: new Matrix2D(1D, 0D, 0.25D, 1D, 16D, 10D));

        Assert.False(reviewed.Matches(current));
    }

    [Fact]
    public void SurvivorIdentityIncludesRelativeNonTextPaintOrder() {
        PdfRedactionTextObjectScope reviewed = CreateReviewedScope(
            owner: 10,
            paintOrderContext: _ => new PdfRedactionPaintOrderContext(1, 2));
        PdfRedactionTextObjectScope current = CreateSurvivorScope(
            owner: 20,
            paintOrderContext: _ => new PdfRedactionPaintOrderContext(2, 2));

        Assert.False(reviewed.Matches(current));
    }

    [Fact]
    public void ReviewedScopesMatchCurrentScopesOneToOne() {
        PdfRedactionTextObjectScope first = CreateReviewedScope(owner: 10);
        PdfRedactionTextObjectScope second = CreateReviewedScope(owner: 11);
        PdfRedactionTextObjectScope current = CreateSurvivorScope(owner: 20);

        int[] matches = PdfRedactionPlan.MatchReviewedTextObjectScopes(
            new[] { first, second },
            new[] { current });

        Assert.Equal(new[] { 0, -1 }, matches);
    }

    private static PdfRedactionTextObjectScope CreateReviewedScope(
        int owner,
        Func<double, PdfRedactionPaintOrderContext>? paintOrderContext = null) {
        PdfContentOrderKey key = PdfContentOrderKey.Root.Append(owner);
        PdfTextSpan span = CreateSpan(
            "AB",
            owner,
            key,
            new[] { 6D, 6D },
            new[] { 1, 1 },
            new[] { new byte[] { 0x41 }, new byte[] { 0x42 } },
            PdfPageClipPath.Rectangle(0D, 0D, 100D, 100D),
            new Matrix2D(1D, 0D, 0D, 1D, 10D, 10D));
        return new PdfRedactionTextObjectScope(
            key,
            new[] { span },
            new[] { new PdfRedactionArea(1, 10D, 0D, 5D, 30D, "remove A") },
            paintOrderContext);
    }

    private static PdfRedactionTextObjectScope CreateSurvivorScope(
        int owner,
        byte glyphByte = 0x42,
        PdfPageClipPath? clipPath = null,
        Matrix2D? transform = null,
        Func<double, PdfRedactionPaintOrderContext>? paintOrderContext = null) {
        PdfContentOrderKey key = PdfContentOrderKey.Root.Append(owner);
        PdfTextSpan span = CreateSpan(
            "B",
            owner,
            key,
            new[] { 6D },
            new[] { 1 },
            new[] { new[] { glyphByte } },
            clipPath ?? PdfPageClipPath.Rectangle(0D, 0D, 100D, 100D),
            transform ?? new Matrix2D(1D, 0D, 0D, 1D, 16D, 10D),
            x: 16D);
        return new PdfRedactionTextObjectScope(key, new[] { span }, paintOrderContext: paintOrderContext);
    }

    private static PdfTextSpan CreateSpan(
        string text,
        int owner,
        PdfContentOrderKey key,
        IReadOnlyList<double> advances,
        IReadOnlyList<int> glyphCharacterLengths,
        IReadOnlyList<byte[]> glyphBytes,
        PdfPageClipPath clipPath,
        Matrix2D transform,
        double x = 10D) =>
        new PdfTextSpan(
            text,
            "F1",
            12D,
            x,
            10D,
            advances.Sum(),
            color: null,
            isVisible: true,
            rotationDegrees: 0D,
            baseFont: "Helvetica",
            clipPath,
            contentOrderKey: key,
            characterAdvances: advances,
            contentStreamObjectNumber: owner,
            textObjectOrderKey: key,
            textToPageTransform: transform,
            glyphCharacterLengths: glyphCharacterLengths,
            glyphBytes: glyphBytes);
}
