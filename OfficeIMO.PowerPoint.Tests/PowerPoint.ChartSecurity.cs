using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

[Collection(PowerPointNonParallelCollection.Name)]
public class PowerPointChartSecurityTests {
    [Fact]
    public void BubbleSnapshot_RejectsOversizedNestedWorkbookPart() {
        using var packageStream = new MemoryStream();
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create(packageStream);
        PowerPointChart chart = presentation.AddSlide().AddChart(
            OfficeChartKind.Bubble,
            CreateBubbleData(pointCount: 1));
        ChartPart chartPart = presentation.Slides[0].SlidePart.ChartParts.Single();
        C.Chart nativeChart = chartPart.ChartSpace!
            .GetFirstChild<C.Chart>()!;
        nativeChart.AddChild(new C.PlotVisibleOnly { Val = true }, true);

        Assert.True(chart.TryGetOfficeSnapshot(out _));

        EmbeddedPackagePart embedded =
            Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
        byte[] oversizedWorkbook = CreateOversizedNestedWorkbook();
        using (Stream target = embedded.GetStream(
                   FileMode.Create,
                   FileAccess.Write)) {
            target.Write(oversizedWorkbook, 0, oversizedWorkbook.Length);
        }

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void BubbleSnapshot_MalformedThemeFailsClosed() {
        string path = Path.Combine(
            Path.GetTempPath(),
            Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create(path)) {
                PowerPointChart chart = presentation.AddSlide().AddChart(
                    OfficeChartKind.Bubble,
                    CreateBubbleData(pointCount: 1));
                Assert.True(chart.TryGetOfficeSnapshot(out _));
                presentation.Save();
            }

            using (var file = new FileStream(
                       path,
                       FileMode.Open,
                       FileAccess.ReadWrite,
                       FileShare.None))
            using (var archive = new ZipArchive(
                       file,
                       ZipArchiveMode.Update,
                       leaveOpen: false)) {
                ZipArchiveEntry theme = archive.Entries.Single(entry =>
                    entry.FullName.StartsWith(
                        "ppt/theme/theme",
                        StringComparison.Ordinal) &&
                    entry.FullName.EndsWith(
                        ".xml",
                        StringComparison.Ordinal));
                string themePath = theme.FullName;
                theme.Delete();
                using StreamWriter writer = new StreamWriter(
                    archive.CreateEntry(themePath).Open(),
                    new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                writer.Write("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><");
            }

            using PowerPointPresentation reopened = PowerPointPresentation.Load(
                path,
                new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly
                });
            PowerPointChart reopenedChart =
                Assert.Single(reopened.Slides[0].Charts);

            Assert.False(reopenedChart.TryGetOfficeSnapshot(out _));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void BubblePointColorMaterialization_IsBoundedAtLargeSupportedScale() {
        const int pointCount = 20_000;
        OfficeChartData data = CreateBubbleData(pointCount);
        var stopwatch = Stopwatch.StartNew();

        using PowerPointPresentation presentation =
            PowerPointPresentation.Create(new MemoryStream());
        presentation.AddSlide().AddChart(OfficeChartKind.Bubble, data);
        stopwatch.Stop();

        C.BubbleChartSeries nativeSeries = presentation.Slides[0].SlidePart
            .ChartParts.Single().ChartSpace!
            .Descendants<C.BubbleChartSeries>().Single();
        Assert.Equal(pointCount, nativeSeries.Elements<C.DataPoint>().Count());
        Assert.True(
            stopwatch.Elapsed < TimeSpan.FromSeconds(30),
            $"Point-color materialization took {stopwatch.Elapsed}.");
    }

    private static OfficeChartData CreateBubbleData(int pointCount) {
        double[] points = Enumerable.Range(1, pointCount)
            .Select(index => (double)index)
            .ToArray();
        OfficeColor?[] colors = Enumerable.Repeat<OfficeColor?>(
            OfficeColor.Parse("#336699"),
            pointCount).ToArray();
        return new OfficeChartData(
            points.Select(value => value.ToString(
                System.Globalization.CultureInfo.InvariantCulture)),
            new[] {
                OfficeChartSeries.CreateBubble(
                    "Series",
                    points,
                    points,
                    points,
                    pointColors: colors)
            });
    }

    private static byte[] CreateOversizedNestedWorkbook() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(
                   stream,
                   ZipArchiveMode.Create,
                   leaveOpen: true)) {
            ZipArchiveEntry worksheet = archive.CreateEntry(
                "xl/worksheets/sheet1.xml",
                CompressionLevel.Optimal);
            using var writer = new StreamWriter(
                worksheet.Open(),
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write(new string(' ', 2 * 1024 * 1024));
            writer.Write("</worksheet>");
        }
        return stream.ToArray();
    }
}
