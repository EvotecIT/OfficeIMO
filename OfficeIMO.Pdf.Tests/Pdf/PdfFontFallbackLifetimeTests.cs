using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using System.Runtime.CompilerServices;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFontFallbackLifetimeTests {
    [Fact]
    public void ResettingAutomaticNamedFallbackReleasesFontBytesAndParsedProgram() {
        (PdfOptions Options, WeakReference[] ReleasedOwners) references = RegisterAndResetAutomaticFallback();
        for (int attempt = 0; attempt < 10 && references.ReleasedOwners.Any(reference => reference.IsAlive); attempt++) {
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            GC.WaitForPendingFinalizers();
        }
        Assert.All(references.ReleasedOwners, reference => Assert.False(reference.IsAlive));
        GC.KeepAlive(references.Options);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static (PdfOptions Options, WeakReference[] ReleasedOwners) RegisterAndResetAutomaticFallback() {
        var options = new PdfOptions();
        options.RegisterAutomaticFontFallbackCandidates(new[] {
            new PdfEmbeddedFontFallbackCandidate("Automatic lifetime", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A'))
        }, new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier });
        string name = Assert.Single(options.EmbeddedFontFallbacks!.FontFamilyNames);
        PdfEmbeddedFontFamily registered = options.NamedFontFamilies[name];
        Assert.True(options.TryResolveNamedFontFace(name, false, false, out PdfNamedFontFace face));
        Assert.True(options.TryGetNamedFontProgramForGeneration(face, out PdfTrueTypeFontProgram? program));
        var references = new[] { new WeakReference(registered!), new WeakReference(registered!.RegularSnapshot),
            new WeakReference(program!), new WeakReference(program!.FontDataForInspection) };
        options.EmbeddedFontFallbacks = null;
        return (options, references);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FallbackNameWhitespaceDoesNotRetainPreviousFontPrograms(bool cff) {
        (byte[] Caller, WeakReference[] ReleasedOwners) references = PlanFallbackNameVariants(cff);
        for (int attempt = 0; attempt < 10 && references.ReleasedOwners.Any(reference => reference.IsAlive); attempt++) {
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            GC.WaitForPendingFinalizers();
        }
        Assert.All(references.ReleasedOwners, reference => Assert.False(reference.IsAlive));
        GC.KeepAlive(references.Caller);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static (byte[] Caller, WeakReference[] ReleasedOwners) PlanFallbackNameVariants(bool cff) {
        string? path = cff
            ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont()
            : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] input = File.ReadAllBytes(path!);
        var first = new PdfEmbeddedFontFallbackCandidate("Lifetime face", input);
        var fallbackSet = new PdfEmbeddedFontFallbackSet(new[] { first });
        PdfTextFallbackPlan plan = fallbackSet.PlanText("A");
        Assert.Equal("A", string.Concat(plan.Segments.Select(segment => segment.Text)));
        PdfEmbeddedFontFallbackCandidate registered = fallbackSet.Candidates[0];
        object program;
        byte[] parsedBytes;
        if (cff) {
            PdfOpenTypeCffFontProgram parsed = PdfFontProgramCache.GetOpenTypeCff(registered.DataSnapshot, registered.FontName);
            program = parsed;
            parsedBytes = parsed.FontDataForInspection;
        } else {
            PdfTrueTypeFontProgram parsed = PdfFontProgramCache.GetTrueType(registered.DataSnapshot, registered.FontName);
            program = parsed;
            parsedBytes = parsed.FontDataForInspection;
        }
        var references = new[] { new WeakReference(first), new WeakReference(registered),
            new WeakReference(first.DataSnapshot), new WeakReference(registered.DataSnapshot),
            new WeakReference(program), new WeakReference(parsedBytes) };
        var next = new PdfEmbeddedFontFallbackCandidate(" Lifetime face", input);
        PdfTextFallbackPlan nextPlan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText("B", new[] { next });
        Assert.Equal("B", string.Concat(nextPlan.Segments.Select(segment => segment.Text)));
        return (input, references);
    }
}
