using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Malformed or unusually deep files must fail with typed rejections or render what they can, never hang or crash.</summary>
public sealed class PdfMalformedInputHardeningTests {
    private const string HelveticaResources = "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>";

    [Fact]
    public void Load_RejectsStreamKeywordAtEndOfInputWithoutLooping() {
        byte[] pdf = Encoding.ASCII.GetBytes("%PDF-1.7\n1 0 obj\ntrailer\n<< /Root 1 0 R /Length 1 0 R >>\nstream");
        using var stop = new CancellationTokenSource(TimeSpan.FromMinutes(1));

        Assert.Throws<PdfParseException>(() => PdfDocument.Load(new MemoryStream(pdf), null, stop.Token).Pages.Count);
    }

    [Fact]
    public void Render_FollowsPageTreeParentCycleOnce() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 /Parent 2 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            Stream("BT /F1 12 Tf 10 100 Td (Loop) Tj ET"));

        PdfPageRenderResult result = Render(pdf);

        Assert.True(result.Succeeded);
    }

    [Fact]
    public void ExtractPages_ValidatesLongOutlineChainsWithoutRecursion() {
        const int items = 20_000;
        const int firstItem = 6;
        var objects = new List<string?> {
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources " + HelveticaResources + " /Contents 4 0 R >>",
            Stream("BT /F1 12 Tf 10 100 Td (Outlined) Tj ET"),
            "<< /Type /Outlines /First " + firstItem + " 0 R /Last " + (firstItem + items - 1) + " 0 R /Count " + items + " >>"
        };
        for (int index = 0; index < items; index++) {
            int number = firstItem + index;
            objects.Add("<< /Title (Item " + index.ToString(CultureInfo.InvariantCulture) + ") /Parent 5 0 R /Dest [3 0 R /Fit]" +
                (index > 0 ? " /Prev " + (number - 1) + " 0 R" : string.Empty) +
                (index < items - 1 ? " /Next " + (number + 1) + " 0 R" : string.Empty) + " >>");
        }

        PdfDocument extracted = PdfDocument.Load(new MemoryStream(BuildPdf(objects))).Pages.Extract("1");

        Assert.Single(PdfDocument.Load(new MemoryStream(extracted.ToBytes())).Pages);
    }

    [Fact]
    public void Render_SkipsTransparencyGroupsNestedBeyondTheThreadStack() {
        // Real papers nest a few dozen isolated groups; each level takes tens of kilobytes of stack in the projection.
        const int depth = 100;
        const int firstForm = 5;
        var objects = new List<string?> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /X " + firstForm + " 0 R >> >> /Contents 4 0 R >>",
            Stream("0 0 1 rg 0 0 50 50 re f /X Do")
        };
        for (int level = 0; level < depth; level++) {
            bool innermost = level == depth - 1;
            string content = innermost ? "1 0 0 rg 20 20 100 100 re f" : "/X Do";
            string resources = innermost ? "<< >>" : "<< /XObject << /X " + (firstForm + level + 1) + " 0 R >> >>";
            objects.Add("<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Group << /S /Transparency /I true >> /Resources " +
                resources + " /Length " + content.Length + " >>\nstream\n" + content + "\nendstream");
        }
        byte[] pdf = BuildPdf(objects);

        PdfPageRenderResult result = RunOnThreadWithStack(1024 * 1024, () => Render(pdf));

        Assert.True(result.Succeeded);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.form-nesting-stack-limit");
    }

    [Fact]
    public void ExtractTextAndRender_SkipStrayCharactersInContentHexStrings() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources " + HelveticaResources + " /Contents 4 0 R >>",
            Stream("BT /F1 12 Tf 10 100 Td <48 65 zz 6C 6C 6F> Tj ET"));

        Assert.Contains("Hello", PdfDocument.Load(new MemoryStream(pdf)).ExtractText(), StringComparison.Ordinal);
        Assert.True(Render(pdf).Succeeded);
    }

    [Fact]
    public void ExtractText_TreatsUnpairedSurrogateFromToUnicodeAsText() {
        const string cmap = "/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CMapName /Lone def " +
            "1 begincodespacerange <00> <FF> endcodespacerange 1 beginbfchar <41> <D800> endbfchar endcmap " +
            "CMapName currentdict /CMap defineresource pop end end";
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            Stream("BT /F1 12 Tf 10 100 Td (A) Tj (Bc) Tj ET"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>",
            Stream(cmap));

        string text = PdfDocument.Load(new MemoryStream(pdf)).ExtractText();

        Assert.Contains("Bc", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_DrawsPagesWithZeroAndNegativeFontSizes() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources " + HelveticaResources + " /Contents 4 0 R >>",
            Stream("BT /F1 0 Tf 10 150 Td (Zero) Tj ET BT /F1 -12 Tf 10 120 Td (Negative) Tj ET BT /F1 12 Tf 10 100 Td (Visible) Tj ET"));

        Assert.True(Render(pdf).Succeeded);
    }

    [Fact]
    public void Render_SkipsAnnotationAppearanceWithoutBoundingBoxAndKeepsThePage() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [5 0 R] /Contents 4 0 R >>",
            Stream("0 0 1 rg 0 0 50 50 re f"),
            "<< /Type /Annot /Subtype /Square /Rect [10 10 100 100] /AP << /N 6 0 R >> >>",
            "<< /Type /XObject /Subtype /Form /Length 21 >>\nstream\n1 0 0 rg 0 0 9 9 re f\nendstream");

        PdfPageRenderResult result = Render(pdf);

        Assert.True(result.Succeeded);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.annotation-appearance-missing");
    }

    [Fact]
    public void Render_OmitsAnElementBeyondTheRasterWorkLimitAndKeepsThePage() {
        // A zig-zag whose vertices share one pixel row needs vertices x row boundaries crossings in that row.
        var path = new StringBuilder("1 0 0 rg 10 100 m ");
        for (int index = 1; index <= 5000; index++) {
            double x = index % 2 == 0 ? 10D : 190D;
            double y = 100D + index * 0.0001D;
            path.Append(x.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(y.ToString("0.0000", CultureInfo.InvariantCulture)).Append(" l ");
        }
        path.Append("h f 0 0 1 rg 0 0 50 50 re f");
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            Stream(path.ToString()));

        PdfPageRenderResult result = Render(pdf);

        Assert.True(result.Succeeded);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.raster.element-work-limit");
    }

    [Fact]
    public void ExtractPages_BlocksStructureTreesTheOrganizerCannotRetarget() {
        byte[] directElement = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 5 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources " + HelveticaResources + " /Contents 4 0 R >>",
            Stream("BT /F1 12 Tf 10 100 Td (Tagged) Tj ET"),
            "<< /Type /StructTreeRoot /K << /Type /StructElem /S /P /Pg 3 0 R /K 0 >> >>");

        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Load(new MemoryStream(directElement)).Pages.Extract("1"));
    }

    [Fact]
    public void ExtractPages_KeepsTaggedPagesWhenTheTrailerInfoObjectIsMissing() {
        byte[] pdf = BuildPdf(new[] {
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 5 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources " + HelveticaResources + " /Contents 4 0 R /StructParents 0 >>",
            Stream("/P << /MCID 0 >> BDC BT /F1 12 Tf 10 100 Td (Tagged) Tj ET EMC"),
            "<< /Type /StructTreeRoot /K 6 0 R >>",
            "<< /Type /StructElem /S /P /P 5 0 R /Pg 3 0 R /K 0 >>"
        }, " /Info 9 0 R");

        PdfDocument extracted = PdfDocument.Load(new MemoryStream(pdf)).Pages.Extract("1");

        Assert.Contains("Tagged", PdfDocument.Load(new MemoryStream(extracted.ToBytes())).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PdfReadLimitKind.ObjectCharacters)]
    [InlineData(PdfReadLimitKind.ObjectTokens)]
    public void ReadLimits_DefaultsAcceptLargeObjectsThatTheUntrustedProfileRejects(PdfReadLimitKind kind) {
        // One 1.5 million character token, like a large inline image or font program written as a string, or a
        // structure parent tree with 40,000 entries in one /Nums array (120,000 tokens).
        string largeObject = kind == PdfReadLimitKind.ObjectCharacters
            ? "(" + new string('A', 1_500_000) + ")"
            : "<< /Nums [" + string.Concat(Enumerable.Range(0, 40_000).Select(static index =>
                index.ToString(CultureInfo.InvariantCulture) + " 4 0 R ")) + "] >>";
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources " + HelveticaResources + " /Contents 4 0 R /PieceInfo << /Data 5 0 R >> >>",
            Stream("BT /F1 12 Tf 10 100 Td (Large) Tj ET"),
            largeObject);

        Assert.Contains("Large", PdfDocument.Load(new MemoryStream(pdf)).ExtractText(), StringComparison.Ordinal);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(new MemoryStream(pdf), PdfLoadOptions.UntrustedDefaults).Pages.Count);
        Assert.Equal(kind, exception.Kind);
    }

    private static PdfPageRenderResult Render(byte[] pdf) =>
        PdfDocument.Load(new MemoryStream(pdf)).Render.Pages("1", new PdfPageRenderOptions { Dpi = 72, Format = OfficeImageExportFormat.Png, ContinueOnError = false })[0];

    private static T RunOnThreadWithStack<T>(int stackBytes, Func<T> action) {
        T result = default!;
        Exception? failure = null;
        var thread = new Thread(() => {
            try {
                result = action();
            } catch (Exception exception) {
                failure = exception;
            }
        }, stackBytes);
        thread.Start();
        thread.Join();
        if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        return result;
    }

    private static string Stream(string content) =>
        "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream";

    private static byte[] BuildPdf(params string[] objects) => BuildPdf((IReadOnlyList<string?>)objects);

    private static byte[] BuildPdf(IReadOnlyList<string?> objects, string trailerEntries = "") {
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new int[objects.Count];
        for (int index = 0; index < objects.Count; index++) {
            if (objects[index] is not string value) continue;
            offsets[index] = pdf.Length;
            pdf.Append(index + 1).Append(" 0 obj\n").Append(value).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        for (int index = 0; index < objects.Count; index++) {
            pdf.Append(objects[index] is null
                ? "0000000000 65535 f \n"
                : offsets[index].ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
        }
        pdf.Append("trailer << /Root 1 0 R /Size ").Append(objects.Count + 1).Append(trailerEntries).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
