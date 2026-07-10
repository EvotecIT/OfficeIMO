using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_ActivatesDataUriFontFacesForLayoutAndImageBackends() {
        byte[] fontData = CreateHtmlRenderTestFont();
        string html = "<style>"
            + "@font-face{font-family:'Scoped Demo';src:url(\"data:font/ttf;base64," + Convert.ToBase64String(fontData) + "\") format('truetype');}"
            + ".scoped{font-family:'Scoped Demo',sans-serif;font-size:100px;line-height:1}"
            + ".fallback{font-family:Arial,sans-serif;font-size:100px;line-height:1}"
            + "</style><p style='margin:0'><span class='scoped'>AA</span><span class='fallback'>BB</span></p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 500D,
            Margins = HtmlRenderMargins.All(8D)
        });

        HtmlRenderText scoped = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "AA");
        HtmlRenderText fallback = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "BB");
        Assert.Single(rendered.Fonts.Faces);
        Assert.Single(rendered.Pages[0].Fonts.Faces);
        Assert.Equal(100D, fallback.X - scoped.X, 6);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontFormatUnsupported);

        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        Assert.Single(drawing.Fonts.Faces);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("@font-face{font-family:\"Scoped Demo\"", svg, StringComparison.Ordinal);
        Assert.Contains(Convert.ToBase64String(fontData), svg, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlRenderAsync_ResolvesAndActivatesExternalFontFacesRelativeToStylesheets() {
        byte[] fontData = CreateHtmlRenderTestFont();
        var requested = new List<string>();
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) => {
                requested.Add(request.Uri.AbsoluteUri);
                if (request.Uri.AbsoluteUri == "https://assets.example.test/css/site.css") {
                    const string css = "@font-face{font-family:'Remote Demo';src:url('../fonts/demo.ttf') format('truetype');}.remote{font-family:'Remote Demo',sans-serif;font-size:100px;line-height:1}";
                    return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css"));
                }

                if (request.Uri.AbsoluteUri == "https://assets.example.test/fonts/demo.ttf") {
                    return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(fontData, "font/ttf"));
                }

                return Task.FromResult<HtmlResolvedResource?>(null);
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/css/site.css'><p class='remote'>AA</p>",
            options);

        Assert.Equal(new[] {
            "https://assets.example.test/css/site.css",
            "https://assets.example.test/fonts/demo.ttf"
        }, requested);
        OfficeFontFace face = Assert.Single(rendered.Fonts.Faces);
        Assert.Equal("Remote Demo", face.FamilyName);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "AA" && text.Font.FamilyName.Contains("Remote Demo", StringComparison.Ordinal));
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
    }

    [Fact]
    public async Task HtmlRenderAsync_DiagnosesUnsupportedWebFontFormatsWithoutAddingCodecs() {
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                new HtmlResolvedResource(new byte[] { 0x77, 0x4F, 0x46, 0x32, 1, 2, 3, 4 }, "font/woff2"))
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(
            "<style>@font-face{font-family:Unsupported;src:url('https://assets.example.test/font.woff2') format('woff2')}p{font-family:Unsupported}</style><p>Fallback</p>",
            options);

        Assert.Empty(rendered.Fonts.Faces);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontFormatUnsupported);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
    }

    [Fact]
    public void HtmlRender_RejectsOversizedFontDataBeforeActivationAndBoundsItsDiagnostic() {
        string data = Convert.ToBase64String(CreateHtmlRenderTestFont());
        string html = "<style>@font-face{font-family:TooLarge;src:url(\"data:font/ttf;base64,"
            + data
            + "\")}p{font-family:TooLarge}</style><p>Fallback</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            MaxResourceBytes = 32L,
            MaxTotalResourceBytes = 32L
        });

        Assert.Empty(rendered.Fonts.Faces);
        HtmlDiagnostic diagnostic = Assert.Single(
            rendered.Diagnostics.Diagnostics,
            item => item.Code == HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded);
        Assert.NotNull(diagnostic.Source);
        Assert.True(diagnostic.Source!.Length < 256);
        Assert.Contains("data:font/ttf;base64,...", diagnostic.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_RenderedProfile_EmbedsAnActiveWebFontFaceWhenAPlatformTtfIsAvailable() {
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? fontPath);
        if (font == null
            || string.IsNullOrWhiteSpace(fontPath)
            || !string.Equals(Path.GetExtension(fontPath), ".ttf", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath!);
        if (fontData.LongLength > 10L * 1024L * 1024L) {
            return;
        }

        string html = "<style>@font-face{font-family:'Pdf Web Demo';src:url(\"data:font/ttf;base64,"
            + Convert.ToBase64String(fontData)
            + "\") format('truetype')}p{font-family:'Pdf Web Demo',sans-serif}</style><p>EmbeddedWebFontMarker</p>";
        HtmlPdfSaveOptions options = HtmlPdfSaveOptions.CreateRenderedProfile();

        byte[] pdf = html.SaveAsPdf(options);

        Assert.Contains("EmbeddedWebFontMarker", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(PdfCore.PdfDiagnostics.Analyze(pdf).EmbeddedFontCount > 0);
        Assert.DoesNotContain(options.RenderDiagnostics!.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
    }

    private static byte[] CreateHtmlRenderTestFont() {
        byte[] cmap = CreateHtmlRenderFormat12Cmap(0x1F600);
        var tables = new List<(string Tag, byte[] Data)> {
            ("cmap", cmap),
            ("glyf", new byte[4]),
            ("head", CreateHtmlRenderHeadTable()),
            ("hhea", CreateHtmlRenderHheaTable()),
            ("hmtx", new byte[] { 0x01, 0xF4, 0x00, 0x00 }),
            ("loca", new byte[4]),
            ("maxp", new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x02 })
        };
        int directoryLength = 12 + tables.Count * 16;
        var offsets = new int[tables.Count];
        int length = directoryLength;
        for (int index = 0; index < tables.Count; index++) {
            offsets[index] = length;
            length += (tables[index].Data.Length + 3) & ~3;
        }

        var font = new byte[length];
        WriteHtmlRenderUInt32(font, 0, 0x00010000);
        WriteHtmlRenderUInt16(font, 4, tables.Count);
        for (int index = 0; index < tables.Count; index++) {
            int record = 12 + index * 16;
            for (int character = 0; character < 4; character++) font[record + character] = (byte)tables[index].Tag[character];
            WriteHtmlRenderUInt32(font, record + 8, (uint)offsets[index]);
            WriteHtmlRenderUInt32(font, record + 12, (uint)tables[index].Data.Length);
            Array.Copy(tables[index].Data, 0, font, offsets[index], tables[index].Data.Length);
        }

        return font;
    }

    private static byte[] CreateHtmlRenderFormat12Cmap(int scalar) {
        var data = new byte[40];
        WriteHtmlRenderUInt16(data, 2, 1);
        WriteHtmlRenderUInt16(data, 4, 3);
        WriteHtmlRenderUInt16(data, 6, 10);
        WriteHtmlRenderUInt32(data, 8, 12);
        WriteHtmlRenderUInt16(data, 12, 12);
        WriteHtmlRenderUInt32(data, 16, 28);
        WriteHtmlRenderUInt32(data, 24, 1);
        WriteHtmlRenderUInt32(data, 28, (uint)scalar);
        WriteHtmlRenderUInt32(data, 32, (uint)scalar);
        WriteHtmlRenderUInt32(data, 36, 1);
        return data;
    }

    private static byte[] CreateHtmlRenderHeadTable() {
        var table = new byte[54];
        WriteHtmlRenderUInt16(table, 18, 1000);
        return table;
    }

    private static byte[] CreateHtmlRenderHheaTable() {
        var table = new byte[36];
        WriteHtmlRenderUInt16(table, 4, 800);
        WriteHtmlRenderUInt16(table, 6, unchecked((ushort)-200));
        WriteHtmlRenderUInt16(table, 34, 1);
        return table;
    }

    private static void WriteHtmlRenderUInt16(byte[] data, int offset, int value) {
        data[offset] = (byte)((value >> 8) & 0xFF);
        data[offset + 1] = (byte)(value & 0xFF);
    }

    private static void WriteHtmlRenderUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)((value >> 24) & 0xFF);
        data[offset + 1] = (byte)((value >> 16) & 0xFF);
        data[offset + 2] = (byte)((value >> 8) & 0xFF);
        data[offset + 3] = (byte)(value & 0xFF);
    }
}
