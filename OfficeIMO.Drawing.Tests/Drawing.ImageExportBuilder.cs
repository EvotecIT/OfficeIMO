using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageExportBuilder_UsesOneFormatAndTerminalGrammar() {
        var options = new TestImageExportOptions();
        var builder = new TestImageExportBuilder(options);

        byte[] png = builder
            .AsPng()
            .WithScale(1.25D)
            .ToBytes();
        byte[] svg = builder
            .AsSvg()
            .WithBackground(OfficeColor.White)
            .ToBytes();

        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
        Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
        Assert.Equal(1.25D, options.Scale);
        Assert.Equal(OfficeColor.White, options.BackgroundColor);
    }

    [Fact]
    public void OfficeImageExportBuilder_SavesTheConfiguredFormat() {
        var builder = new TestImageExportBuilder(new TestImageExportOptions());
        using var png = new MemoryStream();
        using var svg = new MemoryStream();

        OfficeImageExportResult pngResult = builder
            .AsPng()
            .Save(png);
        OfficeImageExportResult svgResult = builder.AsSvg().Save(svg);

        Assert.Equal(OfficeImageExportFormat.Png, pngResult.Format);
        Assert.Equal(OfficeImageExportFormat.Svg, svgResult.Format);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.ToArray().Take(4).ToArray());
        Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg.ToArray()), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void OfficeImageExportBuilder_ExposesSharedRasterFormats(OfficeImageExportFormat format) {
        var options = new TestImageExportOptions();
        var builder = new TestImageExportBuilder(options);

        OfficeImageExportResult result = format switch {
            OfficeImageExportFormat.Jpeg => builder.AsJpeg().Export(),
            OfficeImageExportFormat.Tiff => builder.AsTiff().Export(),
            OfficeImageExportFormat.Webp => builder.AsWebp().Export(),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };

        Assert.Equal(format, result.Format);
        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(result.Bytes).MimeType);
    }

    [Fact]
    public void OfficeImageExportBuilder_ConfiguresRasterEncodingWithoutExposingOptions() {
        var options = new TestImageExportOptions();

        new TestImageExportBuilder(options)
            .WithRasterEncoding(settings => {
                settings.Jpeg.Quality = 73;
                settings.Tiff.Compression = OfficeTiffCompression.None;
            })
            .AsJpeg()
            .Export();

        Assert.Equal(73, options.RasterEncoding.Jpeg.Quality);
        Assert.Equal(OfficeTiffCompression.None, options.RasterEncoding.Tiff.Compression);
    }

    [Fact]
    public void OfficeImageExportBatchBuilder_SavesTheConfiguredFormat() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(new TestImageExportOptions());

            IReadOnlyList<OfficeImageExportResult> pngResults = builder
                .AsPng()
                .Save(folder);
            IReadOnlyList<OfficeImageExportResult> svgResults = builder.AsSvg().Save(folder);

            Assert.Equal(OfficeImageExportFormat.Png, Assert.Single(pngResults).Format);
            Assert.Equal(OfficeImageExportFormat.Svg, Assert.Single(svgResults).Format);
            Assert.True(File.Exists(Path.Combine(folder, "batch.png")));
            Assert.True(File.Exists(Path.Combine(folder, "batch.svg")));
        } finally {
            if (Directory.Exists(folder)) {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg, "batch.jpg")]
    [InlineData(OfficeImageExportFormat.Tiff, "batch.tiff")]
    [InlineData(OfficeImageExportFormat.Webp, "batch.webp")]
    public void OfficeImageExportBatchBuilder_UsesTheConfiguredRasterExtension(
        OfficeImageExportFormat format,
        string expectedFileName) {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(new TestImageExportOptions());

            builder.As(format).Save(folder);

            Assert.True(File.Exists(Path.Combine(folder, expectedFileName)));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void OfficeImageExportBatchBuilder_UsesPortableUniqueFileNames() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBatchBuilder(
                new TestImageExportOptions(),
                "Quarter:1",
                "Quarter?1",
                "CON",
                "COM¹",
                "LPT³.log",
                " trailing. ",
                "A/B");

            builder.Save(folder);

            Assert.True(File.Exists(Path.Combine(folder, "Quarter_1.png")));
            Assert.True(File.Exists(Path.Combine(folder, "Quarter_1-2.png")));
            Assert.True(File.Exists(Path.Combine(folder, "_CON.png")));
            Assert.True(File.Exists(Path.Combine(folder, "_COM¹.png")));
            Assert.True(File.Exists(Path.Combine(folder, "_LPT³.log.png")));
            Assert.True(File.Exists(Path.Combine(folder, "trailing.png")));
            Assert.True(File.Exists(Path.Combine(folder, "A_B.png")));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void OfficeImageExportBuilders_DoNotExposeDuplicateTerminalAliases() {
        string[] removedNames = {
            "AtScale", "OnBackground", "Preview", "HighResolution",
            "ExportPng", "ExportSvg", "ToPng", "ToPngBytes", "ToSvg", "ToSvgBytes",
            "SaveTo", "SaveAsPng", "SavePng", "SaveAsSvg", "SaveSvg"
        };
        string[] singleMethods = typeof(TestImageExportBuilder).GetMethods().Select(method => method.Name).ToArray();
        string[] batchMethods = typeof(TestImageExportBatchBuilder).GetMethods().Select(method => method.Name).ToArray();

        Assert.All(removedNames, name => Assert.DoesNotContain(name, singleMethods));
        Assert.All(removedNames, name => Assert.DoesNotContain(name, batchMethods));
    }

    [Fact]
    public void FileSaveValidatesExtensionAndReturnsNormalizedCommittedPath() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var builder = new TestImageExportBuilder(new TestImageExportOptions()).AsPng();

            Assert.Throws<ArgumentException>(() => builder.Save(Path.Combine(folder, "wrong.jpg")));
            OfficeImageExportResult saved = builder.Save(Path.Combine(folder, "preview"));

            string expected = Path.GetFullPath(Path.Combine(folder, "preview.png"));
            Assert.Equal(expected, saved.SavedPath);
            Assert.True(File.Exists(expected));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void ExportResultUsesTheSameSafeCommitContractAsFluentBuilders() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            var result = new OfficeImageExportResult(
                OfficeImageExportFormat.Png,
                1,
                1,
                OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White)));

            OfficeImageExportResult saved = result.Save(Path.Combine(folder, "direct"));

            Assert.Equal(Path.GetFullPath(Path.Combine(folder, "direct.png")), saved.SavedPath);
            Assert.True(File.Exists(saved.SavedPath));
            Assert.Throws<IOException>(() => result.Save(Path.Combine(folder, "direct.png")));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void FileSaveUsesExplicitSafeConflictPolicies() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(folder, "preview.png");
        try {
            var builder = new TestImageExportBuilder(new TestImageExportOptions()).AsPng();
            builder.Save(path);

            Assert.Throws<IOException>(() => builder.Save(path));
            OfficeImageExportResult unique = builder
                .OnFileConflict(OfficeImageExportFileConflictPolicy.CreateUnique)
                .Save(path);
            Assert.EndsWith("preview-2.png", unique.SavedPath, StringComparison.Ordinal);

            OfficeImageExportResult replaced = builder
                .OnFileConflict(OfficeImageExportFileConflictPolicy.Replace)
                .Save(path);
            Assert.Equal(Path.GetFullPath(path), replaced.SavedPath);
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void BatchSaveFilesReturnsPathsAndMetadataWithoutPayloadContract() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            OfficeImageExportBatchSaveResult saved = new TestImageExportBatchBuilder(
                    new TestImageExportOptions(),
                    "one",
                    "two")
                .SaveFiles(folder);

            Assert.Equal(2, saved.Files.Count);
            Assert.Equal(2, saved.Report.ResultCount);
            Assert.All(saved.Files, file => {
                Assert.True(File.Exists(file.Path));
                Assert.True(file.EncodedLength > 0);
                Assert.Equal(OfficeImageExportFormat.Png, file.Format);
            });
            Assert.DoesNotContain(
                typeof(OfficeImageExportSavedFile).GetProperties(),
                property => property.Name == nameof(OfficeImageExportResult.Bytes));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void BatchBudgetsStopStreamingBeforeUnboundedOutput() {
        var options = new TestImageExportOptions {
            MaximumOutputCount = 1
        };
        var builder = new TestImageExportBatchBuilder(options, "one", "two");
        int consumed = 0;

        OfficeImageExportBatchLimitException exception = Assert.Throws<OfficeImageExportBatchLimitException>(
            () => builder.ExportEach(_ => consumed++));

        Assert.Equal(1, consumed);
        Assert.Equal(nameof(OfficeImageExportOptions.MaximumOutputCount), exception.LimitName);
    }

    [Fact]
    public void GuardedConsumerEnforcesBudgetsForDirectStreamingEntryPoints() {
        var options = new TestImageExportOptions {
            MaximumOutputCount = 1
        };
        int consumed = 0;
        OfficeImageExportConsumer accept =
            OfficeImageExportBatchProcessor.CreateGuardedConsumer(
                options,
                _ => consumed++);
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));

        accept(new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png));
        OfficeImageExportBatchLimitException exception =
            Assert.Throws<OfficeImageExportBatchLimitException>(
                () => accept(new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png)));

        Assert.Equal(1, consumed);
        Assert.Equal(nameof(OfficeImageExportOptions.MaximumOutputCount), exception.LimitName);
    }

    [Fact]
    public void DiagnosticPolicyAggregatesLossAndRejectsConfiguredCodes() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var diagnostics = new[] {
            new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "APPROXIMATED",
                "A feature was approximated.",
                lossKind: OfficeImageExportLossKind.Approximation),
            new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "OMITTED",
                "A feature was omitted.",
                lossKind: OfficeImageExportLossKind.Omission)
        };
        var result = new OfficeImageExportResult(
            OfficeImageExportFormat.Png,
            1,
            1,
            png,
            diagnostics: diagnostics);
        OfficeImageExportReport report = result.CreateReport();

        Assert.True(report.HasLoss);
        Assert.True(report.HasOmissions);
        Assert.Throws<OfficeImageExportPolicyException>(() =>
            result.Require(new OfficeImageExportPolicy { RequireNoOmissions = true }));
        Assert.Throws<OfficeImageExportPolicyException>(() =>
            result.Require(new OfficeImageExportPolicy { FailOnDiagnosticCodes = new[] { "approximated" } }));
    }

    [Fact]
    public void BatchProcessorPreservesOrderWithBoundedParallelRendering() {
        int[] items = Enumerable.Range(1, 8).ToArray();
        var names = new List<string>();

        OfficeImageExportBatchProcessor.ForEachOrdered(
            items,
            maximumDegreeOfParallelism: 3,
            (item, _, token) => {
                token.ThrowIfCancellationRequested();
                return new OfficeImageExportResult(
                    OfficeImageExportFormat.Png,
                    1,
                    1,
                    OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White)),
                    item.ToString());
            },
            result => names.Add(result.Name!));

        Assert.Equal(items.Select(item => item.ToString()), names);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    public void SharedRasterEncodingWritesConsistentPhysicalResolution(OfficeImageExportFormat format) {
        var encoding = new OfficeRasterEncodingOptions {
            DpiX = 144D,
            DpiY = 120D
        };
        byte[] bytes = OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(144, 120, OfficeColor.White),
            format,
            encoding);
        var result = new OfficeImageExportResult(format, 144, 120, bytes);

        Assert.InRange(result.DpiX, 143.98D, 144.02D);
        Assert.InRange(result.DpiY, 119.98D, 120.02D);
        Assert.InRange(result.PhysicalWidthInches, 0.999D, 1.001D);
        Assert.InRange(result.PhysicalHeightInches, 0.999D, 1.001D);
    }

    [Fact]
    public void ExplicitPrintProfileReplacesVagueScalePreset() {
        var options = new TestImageExportOptions();

        OfficeImageExportResult result = new TestImageExportBuilder(options)
            .ForPrint(300D)
            .Export();

        Assert.Equal(300D, options.TargetDpi);
        Assert.Equal(300D / 96D, options.Scale, precision: 8);
        Assert.DoesNotContain(
            typeof(TestImageExportBuilder).GetMethods(),
            method => method.Name == "ForHighResolution");
        Assert.True(result.Width > 300);
    }
}
