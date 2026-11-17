using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfShapedRunReuseTests {
    private const string Text = "Shared café text";

    [Fact]
    public void ReusedShapingReplaysEffectsIntoTheCurrentDocumentWithoutRetainingEarlierRecorders() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "OfficeIMOBaselineSans-Regular.ttf"));
        PdfTrueTypeFontProgram shared = PdfTrueTypeFontProgram.Parse(font, "Reuse");
        var firstCalls = new List<string>();
        WeakReference firstRecorderOwner = ShapeInFirstDocument(shared, firstCalls, out int[] firstGlyphs);

        PdfTrueTypeFontProgram second = shared.ForkForDocument();
        var secondCalls = new List<string>();
        PdfGlyphRun run = second.ShapeText(Text, Options(secondCalls));

        Assert.NotEmpty(firstCalls);
        Assert.Equal(firstCalls, secondCalls);
        Assert.Equal(firstGlyphs, second.GetUsedGlyphIds());
        Assert.NotEmpty(run.Glyphs);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        Assert.False(firstRecorderOwner.IsAlive);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference ShapeInFirstDocument(PdfTrueTypeFontProgram shared, List<string> calls, out int[] glyphs) {
        PdfTrueTypeFontProgram first = shared.ForkForDocument();
        var owner = new RecorderOwner(calls);
        first.ShapeText(Text, PdfTextShapingOptions.ForRendering("Reuse", PdfTextShapingMode.OpenTypeLigatures,
            OfficeManagedTextShapingProvider.Instance, owner.Record));
        glyphs = first.GetUsedGlyphIds().ToArray();
        return new WeakReference(owner);
    }

    private static PdfTextShapingOptions Options(List<string> calls) =>
        PdfTextShapingOptions.ForRendering("Reuse", PdfTextShapingMode.OpenTypeLigatures,
            OfficeManagedTextShapingProvider.Instance, (text, fontName, arg3, arg4) => calls.Add($"{text}|{fontName}|{arg3}|{arg4}"));

    private sealed class RecorderOwner {
        private readonly List<string> _calls;

        internal RecorderOwner(List<string> calls) => _calls = calls;

        internal void Record(string text, string fontName, bool arg3, bool arg4) => _calls.Add($"{text}|{fontName}|{arg3}|{arg4}");
    }
}
