using OfficeIMO.Drawing;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Notebook;
using OfficeIMO.Reader.Subtitles;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMediaAdapterTests {
    [Fact]
    public void ImageAdapter_EmitsMetadataAssetAndOcrCandidateWithoutRunningOcr() {
        byte[] png = CreatePng(3, 2);
        using var stream = new MemoryStream(png, writable: false);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "diagram.png");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/png", asset.MediaType);
        Assert.Equal(3, asset.Width);
        Assert.Equal(2, asset.Height);
        Assert.Equal(png, asset.PayloadBytes);
        Assert.True(asset.PayloadHashMatches(out _));
        Assert.Equal(asset.Id, Assert.Single(result.OcrCandidates).AssetId);
        Assert.Contains("OCR has not been run", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(OfficeDocumentReaderBuilderImageExtensions.HandlerId, result.CapabilitiesUsed);
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void ImageAdapter_CanProjectMetadataWithoutRetainingPayload() {
        byte[] png = CreatePng(1, 1);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddImageHandler(new ReaderImageOptions {
                IncludePayload = false,
                CreateOcrCandidate = false
            })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(png, "pixel.png");

        Assert.Null(Assert.Single(result.Assets).PayloadBytes);
        Assert.Empty(result.OcrCandidates);
        Assert.DoesNotContain("bytes are available", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("source bytes were not retained", result.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void NotebookAdapter_ProjectsMarkdownCodeAndTextOutputsInCellOrder() {
        const string notebook = """
            {
              "cells": [
                { "cell_type": "markdown", "source": ["# Analysis\n", "Intro"] },
                {
                  "cell_type": "code",
                  "source": ["print('ok')"],
                  "outputs": [
                    { "output_type": "stream", "name": "stdout", "text": ["ok\n"] },
                    { "output_type": "display_data", "data": { "text/markdown": ["**done**"] } }
                  ]
                }
              ],
              "metadata": { "kernelspec": { "language": "python" } },
              "nbformat": 4,
              "nbformat_minor": 5
            }
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "analysis.ipynb");

        Assert.Equal(2, result.Chunks.Count);
        Assert.Contains("# Analysis", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("```python", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("print('ok')", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("**done**", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(result.Metadata, item => item.Name == "NbFormat" && item.Value == "4");
        Assert.Contains(OfficeDocumentReaderBuilderNotebookExtensions.HandlerId, result.CapabilitiesUsed);
    }

    [Fact]
    public void NotebookAdapter_ReportsConfiguredCellLimit() {
        const string notebook = """
            {
              "cells": [
                { "cell_type": "markdown", "source": "one" },
                { "cell_type": "markdown", "source": "two" }
              ],
              "metadata": {},
              "nbformat": 4,
              "nbformat_minor": 0
            }
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddNotebookHandler(new ReaderNotebookOptions { MaxCells = 1 })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "bounded.ipynb");

        Assert.Single(result.Chunks);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "notebook-cell-limit");
    }

    [Fact]
    public void SubtitleAdapter_ProjectsSrtCueTextTimingAndSourceLines() {
        const string srt = """
            1
            00:00:01,250 --> 00:00:03,500
            <i>Hello</i> &amp; welcome

            2
            00:00:04,000 --> 00:00:05,000
            Second cue
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "captions.srt");

        Assert.Equal(2, result.Chunks.Count);
        Assert.Equal("Hello & welcome", result.Chunks[0].Text);
        Assert.Contains("00:00:01.250 → 00:00:03.500", result.Chunks[0].Markdown, StringComparison.Ordinal);
        Assert.Equal(1, result.Chunks[0].Location.StartLine);
        OfficeDocumentMetadataEntry timing = Assert.Single(result.Metadata, item => item.Id == "subtitle-cue-000000-timing");
        Assert.Equal("1250", timing.Attributes["startMilliseconds"]);
        Assert.Equal("3500", timing.Attributes["endMilliseconds"]);
        Assert.Equal("captions.srt", timing.Location?.Path);
        Assert.Contains(OfficeDocumentReaderBuilderSubtitleExtensions.HandlerId, result.CapabilitiesUsed);
    }

    [Fact]
    public void SubtitleAdapter_ParsesWebVttAndSkipsMetadataBlocks() {
        const string vtt = """
            WEBVTT

            NOTE generated captions
            ignored note

            cue-a
            00:00:00.500 --> 00:00:02.000 align:start
            <v Speaker>Welcome</v>
            """;
        byte[] bytes = Encoding.UTF8.GetBytes(vtt);
        using var stream = new MemoryStream(bytes, writable: false);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "captions.vtt");

        Assert.Equal("Welcome", Assert.Single(result.Chunks).Text);
        Assert.DoesNotContain("ignored note", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(result.Metadata, item => item.Name == "Format" && item.Value == "webvtt");
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void AllPreset_IncludesDependencyFreeMediaAdapters() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        IReadOnlyList<ReaderHandlerCapability> capabilities = reader.GetCapabilities();

        Assert.Contains(capabilities, item => item.Id == OfficeDocumentReaderBuilderImageExtensions.HandlerId);
        Assert.Contains(capabilities, item => item.Id == OfficeDocumentReaderBuilderNotebookExtensions.HandlerId);
        Assert.Contains(capabilities, item => item.Id == OfficeDocumentReaderBuilderSubtitleExtensions.HandlerId);
    }

    private static byte[] CreatePng(int width, int height) {
        byte[] rgba = new byte[checked(width * height * 4)];
        for (int index = 0; index < rgba.Length; index += 4) {
            rgba[index] = 32;
            rgba[index + 1] = 64;
            rgba[index + 2] = 128;
            rgba[index + 3] = 255;
        }
        return OfficePngWriter.EncodeRgba(width, height, rgba);
    }
}
