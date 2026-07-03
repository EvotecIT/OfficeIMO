using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentVisualBaselineTests {
    [Fact]
    public void RepresentativeReport_MatchesVisualGeometryBaseline() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10.5,
            FooterFont = PdfStandardFont.Helvetica,
            FooterFontSize = 8,
            FooterFormat = "OfficeIMO.Pdf visual gate - page {page}/{pages}",
            FooterAlign = PdfAlign.Right,
            ShowPageNumbers = true
        };

        byte[] bytes = PdfDocument.Create(options)
            .Meta(title: "OfficeIMO.Pdf Visual Baseline", author: "OfficeIMO")
            .H1("OfficeIMO.Pdf Visual Baseline", PdfAlign.Left, PdfColor.FromRgb(32, 76, 120))
            .Paragraph(p => p
                .Text("This report combines ")
                .Bold("rich text")
                .Text(", proportional Helvetica spacing, wrapped table cells, lists, panels, and footer text."))
            .PanelParagraph(
                p => p.Text("Panel content should sit inside the page margins with comfortable padding and without clipping."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(245, 248, 252),
                    BorderColor = PdfColor.FromRgb(150, 170, 190),
                    PaddingX = 8,
                    PaddingY = 7
                })
            .Bullets(new[] {
                "Readable line spacing",
                "Stable text geometry",
                "No right-edge overflow"
            })
            .Table(new[] {
                new[] { "Area", "Expectation", "Status" },
                new[] { "Typography", "Natural word spacing for standard proportional fonts.", "Guarded" },
                new[] { "Tables", "Long content wraps inside cells instead of drawing through adjacent columns or page margins.", "Guarded" },
                new[] { "PowerShell", "PSWriteOffice can eventually expose this engine as the PSWritePDF successor.", "Roadmap" }
            }, style: TableStyles.RightAlignedNumbers())
            .Paragraph(p => p.Text("End of visual baseline."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();

        string snapshot = BuildVisualSnapshot(bytes);
        AssertVisualBaseline("officeimo-pdf-representative-report", snapshot);
    }

    [Fact]
    public void ProfessionalReport_MatchesVisualContentBaseline() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            HeaderFont = PdfStandardFont.Helvetica,
            HeaderFontSize = 8,
            HeaderFormat = "OfficeIMO.Pdf professional report",
            HeaderAlign = PdfAlign.Left,
            HeaderOffsetY = 18,
            ShowHeader = true,
            FooterFont = PdfStandardFont.Helvetica,
            FooterFontSize = 8,
            FooterFormat = "OfficeIMO.Pdf professional report - page {page}/{pages}",
            FooterAlign = PdfAlign.Center,
            ShowPageNumbers = true,
            CreateOutlineFromHeadings = true
        };

        byte[] logo = CreateMinimalJpeg(200, 100);
        byte[] alphaBadge = CreateTransparentBadgePng();

        byte[] bytes = PdfDocument.Create(options)
            .Meta(title: "OfficeIMO.Pdf Professional Report", author: "OfficeIMO")
            .H1("Executive Security Summary", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p
                .Text("A polished first-party PDF report built with ")
                .Bold("OfficeIMO.Pdf")
                .Text(", shared drawing descriptors, wrapped tables, images, and generated bookmarks."))
            .Image(logo, 120, 60, PdfAlign.Right, fit: OfficeImageFit.Contain)
            .PanelParagraph(
                p => p.Bold("Healthy").Text(" posture with no critical findings."),
                CreateStatusPanelStyle(),
                defaultColor: PdfColor.FromRgb(42, 132, 82))
            .Image(alphaBadge, 18, 18, PdfAlign.Left)
            .Shape(CreateAccentRibbon(), spacingBefore: 12, spacingAfter: 10)
            .PanelParagraph(
                p => p.Text("Operator note: long values should wrap cleanly, tables should stay inside the page, and reusable drawing primitives should remain available to Word, Excel, PowerPoint, and PDF exporters."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                })
            .Table(new[] {
                new[] { "Signal", "Evidence", "Action" },
                new[] { "DMARC", "Policy is enforced for the primary domain and aligned subdomains.", "Monitor" },
                new[] { "TLS", "Certificate chain and protocol posture are ready for automated PSWriteOffice reports.", "Keep" },
                new[] { "DNS", "Delegation and stale-record checks are summarized without overflowing table cells.", "Review" },
                new[] { "PDF", "The report is generated by the MIT licensed dependency-free OfficeIMO.Pdf engine.", "Expand" }
            }, style: CreateReportTableStyle())
            .Image(logo, 112, 36, PdfAlign.Center, fit: OfficeImageFit.Contain)
            .Paragraph(p => p.Text("Generated by OfficeIMO.Pdf."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();

        string snapshot = BuildVisualSnapshot(bytes);
        AssertVisualBaseline("officeimo-pdf-professional-report", snapshot);
    }

    private static string BuildVisualSnapshot(byte[] pdfBytes) {
        using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
        var sb = new StringBuilder();
        sb.AppendLine("pages=" + pdf.NumberOfPages.ToString(CultureInfo.InvariantCulture));

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            var visibleLetters = page.Letters
                .Where(static letter => !string.IsNullOrEmpty(letter.Value) && !char.IsWhiteSpace(letter.Value[0]))
                .ToList();

            double minX = visibleLetters.Count == 0 ? 0 : visibleLetters.Min(static letter => letter.StartBaseLine.X);
            double maxX = visibleLetters.Count == 0 ? 0 : visibleLetters.Max(static letter => letter.EndBaseLine.X);
            double minY = visibleLetters.Count == 0 ? 0 : visibleLetters.Min(static letter => letter.StartBaseLine.Y);
            double maxY = visibleLetters.Count == 0 ? 0 : visibleLetters.Max(static letter => letter.StartBaseLine.Y);

            sb.AppendLine("[page:" + pageNumber.ToString(CultureInfo.InvariantCulture) + "]");
            sb.AppendLine("size=" + Format(page.Width) + "x" + Format(page.Height));
            sb.AppendLine("visibleLetters=" + visibleLetters.Count.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("bounds=x:" + Format(minX) + ".." + Format(maxX) + " y:" + Format(minY) + ".." + Format(maxY));
            sb.AppendLine("rightOverflow=" + Format(Math.Max(0, maxX - (page.Width - 72))));

            var lines = page.Letters
                .Where(static letter => !string.IsNullOrEmpty(letter.Value))
                .GroupBy(static letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(static group => group.Key)
                .Select(static group => {
                    var ordered = group.OrderBy(static letter => letter.StartBaseLine.X).ToList();
                    string text = string.Concat(ordered.Select(static letter => letter.Value)).Trim();
                    double x1 = ordered.Min(static letter => letter.StartBaseLine.X);
                    double x2 = ordered.Max(static letter => letter.EndBaseLine.X);
                    double pointSize = ordered
                        .Where(static letter => !string.IsNullOrWhiteSpace(letter.Value))
                        .Select(static letter => letter.PointSize)
                        .DefaultIfEmpty(0)
                        .Max();

                    return new VisualLine(group.Key, x1, x2, pointSize, NormalizeLineText(text));
                })
                .Where(static line => line.Text.Length > 0)
                .ToList();

            sb.AppendLine("lines=" + lines.Count.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < lines.Count; i++) {
                var line = lines[i];
                sb.AppendLine(
                    "line[" + i.ToString(CultureInfo.InvariantCulture) + "]=" +
                    "y:" + Format(line.Y) +
                    " x:" + Format(line.X1) +
                    " w:" + Format(line.X2 - line.X1) +
                    " size:" + Format(line.PointSize) +
                    " text:" + line.Text);
            }
        }

        AppendContentStreamSignals(sb, pdfBytes);

        return sb.ToString().TrimEnd();
    }

    private static void AppendContentStreamSignals(StringBuilder sb, byte[] pdfBytes) {
        string pdfText = Encoding.ASCII.GetString(pdfBytes);
        string content = ExtractNonImageStreams(pdfText);
        int imageDraws = CountOccurrences(content, " Do");
        int imageSoftMasks = CountOccurrences(pdfText, "/SMask");
        int visibleImageXObjects = Math.Max(imageDraws, CountOccurrences(pdfText, "/Subtype /Image") - imageSoftMasks);

        sb.AppendLine("[content-streams]");
        sb.AppendLine("imageXObjects=" + visibleImageXObjects.ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("imageSoftMasks=" + imageSoftMasks.ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("imageDraws=" + imageDraws.ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("clipOps=" + CountOccurrences(content, " W n").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("graphicsStateResources=" + CountOccurrences(pdfText, "/Type /ExtGState").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("graphicsStateUses=" + CountOccurrences(content, " gs").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("shadingResources=" + CountOccurrences(pdfText, "/ShadingType 2").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("shadingDraws=" + CountOccurrences(content, " sh").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("rectOps=" + CountOccurrences(content, " re").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("curveOps=" + CountOccurrences(content, " c").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("saveOps=" + CountOperatorLines(content, "q").ToString(CultureInfo.InvariantCulture));
        sb.AppendLine("restoreOps=" + CountOperatorLines(content, "Q").ToString(CultureInfo.InvariantCulture));

        var visualOps = content
            .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(NormalizeContentLine)
            .Where(static line => line.Contains("/Im") || line.Contains(" W n") || line.Contains(" sh") || line.Contains(" gs"))
            .Take(20)
            .ToList();

        sb.AppendLine("visualOps=" + visualOps.Count.ToString(CultureInfo.InvariantCulture));
        for (int i = 0; i < visualOps.Count; i++) {
            sb.AppendLine("visualOp[" + i.ToString(CultureInfo.InvariantCulture) + "]=" + visualOps[i]);
        }
    }

    private static string ExtractNonImageStreams(string pdfText) {
        var sb = new StringBuilder();
        int searchIndex = 0;
        while (true) {
            int streamIndex = pdfText.IndexOf("stream\n", searchIndex, StringComparison.Ordinal);
            if (streamIndex < 0) {
                break;
            }

            int endIndex = pdfText.IndexOf("\nendstream", streamIndex, StringComparison.Ordinal);
            if (endIndex < 0) {
                break;
            }

            int objectIndex = pdfText.LastIndexOf(" obj\n", streamIndex, StringComparison.Ordinal);
            string header = objectIndex >= 0
                ? pdfText.Substring(objectIndex, streamIndex - objectIndex)
                : pdfText.Substring(Math.Max(0, streamIndex - 300), Math.Min(300, streamIndex));

            if (!header.Contains("/Subtype /Image")) {
                int contentStart = streamIndex + "stream\n".Length;
                sb.AppendLine(pdfText.Substring(contentStart, endIndex - contentStart));
            }

            searchIndex = endIndex + "\nendstream".Length;
        }

        return sb.ToString();
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static int CountOperatorLines(string content, string operatorName) =>
        content
            .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(static line => line.Trim())
            .Count(line => string.Equals(line, operatorName, StringComparison.Ordinal));

    private static string NormalizeContentLine(string line) {
        var sb = new StringBuilder(line.Length);
        bool previousWhitespace = false;
        foreach (char ch in line.Trim()) {
            if (char.IsWhiteSpace(ch)) {
                if (!previousWhitespace) sb.Append(' ');
                previousWhitespace = true;
                continue;
            }

            sb.Append(ch);
            previousWhitespace = false;
        }

        string normalized = NormalizeImageResourceNames(sb.ToString());
        return normalized.Length <= 160 ? normalized : normalized.Substring(0, 160);
    }

    private static string NormalizeImageResourceNames(string line) {
        var sb = new StringBuilder(line.Length);
        for (int i = 0; i < line.Length; i++) {
            if (i + 3 < line.Length &&
                line[i] == '/' &&
                line[i + 1] == 'I' &&
                line[i + 2] == 'm' &&
                char.IsDigit(line[i + 3])) {
                sb.Append("/Im#");
                i += 3;
                while (i + 1 < line.Length && char.IsDigit(line[i + 1])) {
                    i++;
                }

                continue;
            }

            sb.Append(line[i]);
        }

        return sb.ToString();
    }

    private static void AssertVisualBaseline(string name, string actualSnapshot) {
        string expectedPath = GetExpectedPath(name);
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_PDF_VISUAL_BASELINE"), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
            File.WriteAllText(expectedPath, actualSnapshot + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return;
        }

        if (!File.Exists(expectedPath)) {
            throw new FileNotFoundException(
                "PDF visual baseline missing. Set OFFICEIMO_UPDATE_PDF_VISUAL_BASELINE=1 and re-run this test to generate it.",
                expectedPath);
        }

        string expected = File.ReadAllText(expectedPath, Encoding.UTF8)
            .Replace("\r\n", "\n")
            .TrimEnd();
        string actual = actualSnapshot
            .Replace("\r\n", "\n")
            .TrimEnd();

        Assert.Equal(expected, actual);
    }

    private static string GetExpectedPath(string fixtureName) =>
        Path.Combine(GetTestsProjectRoot(), "Pdf", "VisualBaselines", fixtureName + ".snapshot.txt");

    private static string GetTestsProjectRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Pdf.Tests.csproj"))) {
                return directory.FullName;
            }

            string pdfProjectRoot = Path.Combine(directory.FullName, "OfficeIMO.Pdf.Tests");
            if (File.Exists(Path.Combine(pdfProjectRoot, "OfficeIMO.Pdf.Tests.csproj"))) {
                return pdfProjectRoot;
            }

            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate OfficeIMO PDF test project root from test runtime base directory.");
    }

    private static string Format(double value) =>
        Math.Round(value, 1, MidpointRounding.AwayFromZero).ToString("0.0", CultureInfo.InvariantCulture);

    private static string NormalizeLineText(string text) {
        var sb = new StringBuilder(text.Length);
        bool previousWhitespace = false;
        foreach (char ch in text) {
            if (char.IsWhiteSpace(ch)) {
                if (!previousWhitespace) sb.Append(' ');
                previousWhitespace = true;
                continue;
            }

            sb.Append(ch);
            previousWhitespace = false;
        }

        string normalized = sb.ToString().Trim();
        return normalized.Length <= 120 ? normalized : normalized.Substring(0, 120);
    }

    private static PanelStyle CreateStatusPanelStyle() {
        return new PanelStyle {
            Background = PdfColor.FromRgb(230, 247, 238),
            BorderColor = PdfColor.FromRgb(42, 132, 82),
            BorderWidth = 1.2,
            PaddingX = 8,
            PaddingY = 5,
            MaxWidth = 245
        };
    }

    private static OfficeShape CreateAccentRibbon() {
        var accent = OfficeShape.RoundedRectangle(168, 8, 4);
        accent.FillGradient = OfficeLinearGradient.Horizontal(
            OfficeColor.FromRgb(32, 76, 120),
            OfficeColor.FromRgb(78, 159, 188));
        accent.Shadow = new OfficeShadow(OfficeColor.Black, 0.16, 1.5, 1.5);
        accent.StrokeColor = OfficeColor.FromRgb(32, 76, 120);
        accent.StrokeWidth = 0;
        return accent;
    }

    private static PdfTableStyle CreateReportTableStyle() {
        return new PdfTableStyle {
            HeaderFill = PdfColor.FromRgb(32, 76, 120),
            HeaderTextColor = PdfColor.White,
            TextColor = PdfColor.FromRgb(31, 41, 55),
            RowStripeFill = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(210, 218, 226),
            BorderWidth = 0.5,
            CellPaddingX = 6,
            CellPaddingY = 5,
            AutoFitColumns = true
        };
    }

    private static byte[] CreateMinimalJpeg(int width, int height) {
        return new byte[] {
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            (byte)(height >> 8), (byte)(height & 0xFF),
            (byte)(width >> 8), (byte)(width & 0xFF),
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        };
    }

    private static byte[] CreateTransparentBadgePng() => PdfPngTestImages.CreateRgbaPng(0x2A, 0x84, 0x52, 0xA0);

    private readonly struct VisualLine {
        internal VisualLine(double y, double x1, double x2, double pointSize, string text) {
            Y = y;
            X1 = x1;
            X2 = x2;
            PointSize = pointSize;
            Text = text;
        }

        internal double Y { get; }
        internal double X1 { get; }
        internal double X2 { get; }
        internal double PointSize { get; }
        internal string Text { get; }
    }
}
