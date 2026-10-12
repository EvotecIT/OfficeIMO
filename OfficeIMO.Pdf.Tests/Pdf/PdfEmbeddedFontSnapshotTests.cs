using OfficeIMO.Pdf;
using System.Runtime.CompilerServices;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfEmbeddedFontSnapshotTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedFamilyConstructionReusesParsingAndIsolatesDocumentGlyphs(bool cff) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont()
            : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] caller = File.ReadAllBytes(path!);
        var first = new PdfOptions().RegisterFontFamily(PdfStandardFont.Helvetica,
            new PdfEmbeddedFontFamily("Repeated snapshot", caller));
        var second = new PdfOptions().RegisterFontFamily(PdfStandardFont.Helvetica,
            new PdfEmbeddedFontFamily("Repeated snapshot", caller));

        if (cff) {
            Assert.True(first.TryGetEmbeddedOpenTypeCffFontProgram(PdfStandardFont.Helvetica, out PdfOpenTypeCffFontProgram? a));
            Assert.True(second.TryGetEmbeddedOpenTypeCffFontProgram(PdfStandardFont.Helvetica, out PdfOpenTypeCffFontProgram? b));
            Assert.NotSame(a, b);
            Assert.Same(a!.FontDataForInspection, b!.FontDataForInspection);
            Assert.True(a.TryGetGlyphId('A', out int ga));
            Assert.True(b.TryGetGlyphId('B', out int gb));
            a.RecordGlyphUsage(ga, 'A');
            b.RecordGlyphUsage(gb, 'B');
            Assert.Equal(new[] { ga }, a.GetUsedGlyphIds());
            Assert.Equal(new[] { gb }, b.GetUsedGlyphIds());
        } else {
            Assert.True(first.TryGetEmbeddedStandardFontProgram(PdfStandardFont.Helvetica, out PdfTrueTypeFontProgram? a));
            Assert.True(second.TryGetEmbeddedStandardFontProgram(PdfStandardFont.Helvetica, out PdfTrueTypeFontProgram? b));
            Assert.NotSame(a, b);
            Assert.Same(a!.FontDataForInspection, b!.FontDataForInspection);
            Assert.True(a.TryGetGlyphId('A', out int ga));
            Assert.True(b.TryGetGlyphId('B', out int gb));
            a.RecordGlyphUsage(ga, 'A');
            b.RecordGlyphUsage(gb, 'B');
            Assert.Equal(new[] { ga }, a.GetUsedGlyphIds());
            Assert.Equal(new[] { gb }, b.GetUsedGlyphIds());
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void CallerMutationCreatesANewSnapshotWithoutChangingEarlierFamilies(int face) {
        byte[] caller = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        PdfEmbeddedFontFamily Create() => new("Mutable snapshot", caller,
            face == 1 ? caller : null, face == 2 ? caller : null, face == 3 ? caller : null);
        byte[] Read(PdfEmbeddedFontFamily family) => face switch {
            1 => family.Bold!, 2 => family.Italic!, 3 => family.BoldItalic!, _ => family.Regular
        };
        PdfEmbeddedFontFamily first = Create();
        PdfEmbeddedFontFamily unchanged = Create();
        Assert.Equal(Read(first), Read(unchanged));
        caller[0] = 255;
        PdfEmbeddedFontFamily changed = Create();
        Assert.Equal(0, Read(first)[0]);
        Assert.Equal(0, Read(unchanged)[0]);
        Assert.Equal(255, Read(changed)[0]);
        byte[] returned = Read(changed);
        returned[0] = 17;
        Assert.Equal(255, Read(changed)[0]);
        caller[0] = 0;
        Assert.Equal(Read(first), Read(Create()));
    }

    [Fact]
    public void ConcurrentFamilyConstructionSharesOnlyPrivateImmutableSnapshots() {
        byte[] caller = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        var families = new PdfEmbeddedFontFamily[32];
        System.Threading.Tasks.Parallel.For(0, families.Length,
            index => families[index] = new PdfEmbeddedFontFamily("Concurrent snapshot", caller));
        Assert.NotSame(caller, families[0].RegularSnapshot);
        Assert.All(families, family => Assert.Same(families[0].RegularSnapshot, family.RegularSnapshot));
        caller[0] = 255;
        Assert.All(families, family => Assert.Equal(0, family.Regular[0]));
    }

    [Fact]
    public void SnapshotReuseDoesNotRetainCallerBytesOrTheirParsedFont() {
        WeakReference[] owners = CreateReleasedOwners();
        for (int attempt = 0; attempt < 10 && owners.Any(owner => owner.IsAlive); attempt++) {
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            GC.WaitForPendingFinalizers();
        }
        Assert.All(owners, owner => Assert.False(owner.IsAlive));
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference[] CreateReleasedOwners() {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] caller = File.ReadAllBytes(path!);
        var family = new PdfEmbeddedFontFamily("Released snapshot", caller);
        var options = new PdfOptions().RegisterFontFamily(PdfStandardFont.Helvetica, family);
        Assert.True(options.TryGetEmbeddedStandardFontProgram(PdfStandardFont.Helvetica, out PdfTrueTypeFontProgram? font));
        return new[] { new WeakReference(caller), new WeakReference(family.RegularSnapshot),
            new WeakReference(font!.FontDataForInspection) };
    }
}
