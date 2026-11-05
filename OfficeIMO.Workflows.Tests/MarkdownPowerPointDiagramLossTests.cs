using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

[CollectionDefinition("Markdown diagram process policy", DisableParallelization = true)]
public sealed class MarkdownDiagramProcessPolicyCollection { }

[Collection("Markdown diagram process policy")]
public sealed class MarkdownPowerPointDiagramLossTests {
    [Fact]
    public async Task MermaidSourceSurvivesNativeSaveWithoutExecutingRendererAndStrictSavePreservesDestination() {
        const string rendererVariable = "OFFICEIMO_MARKUP_MERMAID_CLI";
        const string diagram = "flowchart LR\n    Start[Source marker] --> Finish[Destination marker]";
        string directory = Directory.CreateTempSubdirectory("officeimo-markdown-diagram-").FullName;
        string marker = Path.Combine(directory, "renderer-started");
        string? previousRenderer = Environment.GetEnvironmentVariable(rendererVariable);
        try {
            string renderer = Path.Combine(directory, OperatingSystem.IsWindows() ? "renderer.cmd" : "renderer.sh");
            string script = OperatingSystem.IsWindows()
                ? $"@echo off\r\necho started>\"{marker}\"\r\nexit /b 1\r\n"
                : $"#!/bin/sh\nprintf started > '{marker}'\nexit 1\n";
            await File.WriteAllTextAsync(renderer, script);
            if (!OperatingSystem.IsWindows())
                File.SetUnixFileMode(renderer, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            Environment.SetEnvironmentVariable(rendererVariable, renderer);

            using MarkdownDocument source = MarkdownDocument.Parse($"# Diagram marker\n\n```mermaid\n{diagram}\n```");
            string output = Path.Combine(directory, "diagram.pptx");
            OfficeConversionReport report = await source.SaveAsPowerPointAsync(output);
            AssertDiagramLoss(report);
            using (PowerPointDocument deck = PowerPointDocument.Load(output)) {
                Assert.Single(deck.Slides);
                Assert.Empty(deck.Slides[0].Images);
            }
            using (ZipArchive archive = ZipFile.OpenRead(output)) {
                using Stream slide = archive.GetEntry("ppt/slides/slide1.xml")!.Open();
                XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
                string text = string.Join("\n", XDocument.Load(slide).Descendants(drawing + "t").Select(node => node.Value));
                Assert.Contains("Diagram marker", text, StringComparison.Ordinal);
                Assert.Contains(diagram, text, StringComparison.Ordinal);
            }
            Assert.False(File.Exists(marker), "Markdown conversion must not execute a discoverable Mermaid renderer.");

            string rejected = Path.Combine(directory, "strict.pptx");
            byte[] sentinel = "unchanged destination"u8.ToArray();
            await File.WriteAllBytesAsync(rejected, sentinel);
            OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() =>
                source.SaveAsPowerPointAsync(rejected, new MarkdownToPowerPointOptions { FailOnLoss = true }));
            AssertDiagramLoss(error.Report);
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(rejected));
            Assert.False(File.Exists(marker));
        } finally {
            Environment.SetEnvironmentVariable(rendererVariable, previousRenderer);
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void AssertDiagramLoss(OfficeConversionReport report) {
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, finding => finding.Format == "OfficeIMO.Markup"
            && finding.Location == "Diagram" && finding.LossKind == OfficeConversionLossKind.Omission);
    }
}
