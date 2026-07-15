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
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "ocr-needed");
        Assert.DoesNotContain("bytes are available", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("source bytes were not retained", result.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageAdapter_RejectsAnImageExtensionWithoutAnImageSignature() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes("not an image"), "renamed.png"));
    }

    [Fact]
    public void ImageAdapter_IdentifiesWebpFromItsHeader() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(CreateWebpExtendedHeader(3, 2), "image.webp");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/webp", asset.MediaType);
        Assert.Equal(3, asset.Width);
        Assert.Equal(2, asset.Height);
    }

    [Fact]
    public void ImageAdapter_IdentifiesContentVerifiedSvgAfterBomAndComment() {
        const string svg = "\uFEFF<!-- generated --><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"3\" height=\"2\"/>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(svg), "image.svg");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/svg+xml", asset.MediaType);
        Assert.Equal(3, asset.Width);
        Assert.Equal(2, asset.Height);
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
    public void NotebookAdapter_PreservesLeadingWhitespaceInMarkdownCells() {
        const string notebook = """
            {
              "cells": [
                { "cell_type": "markdown", "source": ["    indented code\n", "\n"] }
              ],
              "metadata": {},
              "nbformat": 4,
              "nbformat_minor": 5
            }
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "indented.ipynb");

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.StartsWith("    indented code", chunk.Markdown, StringComparison.Ordinal);
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
    public void NotebookAdapter_WarnsWhenOneOutputFillsTheBudgetBeforeAnotherOutput() {
        const string notebook = """
            {
              "cells": [
                {
                  "cell_type": "code",
                  "source": "print('bounded')",
                  "outputs": [
                    { "output_type": "stream", "text": "1234" },
                    { "output_type": "stream", "text": "omitted" }
                  ]
                }
              ],
              "metadata": {},
              "nbformat": 4,
              "nbformat_minor": 0
            }
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddNotebookHandler(new ReaderNotebookOptions { MaxOutputCharactersPerCell = 4 })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "bounded.ipynb");

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Contains("1234", chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("omitted", chunk.Text, StringComparison.Ordinal);
        Assert.Contains(chunk.Warnings!, warning => warning.Contains("MaxOutputCharactersPerCell", StringComparison.Ordinal));
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
    public void SubtitleAdapter_ParsesHoursBeyondOneDay() {
        const string srt = "1\n24:00:00,000 --> 25:01:02,003\nLong recording\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "long.srt");

        ReaderChunk cue = Assert.Single(result.Chunks);
        Assert.Contains("24:00:00.000 → 25:01:02.003", cue.Markdown, StringComparison.Ordinal);
        OfficeDocumentMetadataEntry timing = Assert.Single(result.Metadata, item => item.Id == "subtitle-cue-000000-timing");
        Assert.Equal("86400000", timing.Attributes["startMilliseconds"]);
        Assert.Equal("90062003", timing.Attributes["endMilliseconds"]);
    }

    [Fact]
    public void SubtitleAdapter_PreservesCueIdentifiersThatOnlyStartWithMetadataWords() {
        const string vtt = """
            WEBVTT

            NOTE1
            00:00:00.000 --> 00:00:01.000
            First cue

            STYLE_intro
            00:00:01.000 --> 00:00:02.000
            Second cue

            REGION-a
            00:00:02.000 --> 00:00:03.000
            Third cue
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(vtt), "identifiers.vtt");

        Assert.Equal(new[] { "First cue", "Second cue", "Third cue" }, result.Chunks.Select(chunk => chunk.Text));
    }

    [Fact]
    public void AllPreset_IncludesDependencyFreeMediaAdapters() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        IReadOnlyList<ReaderHandlerCapability> capabilities = reader.GetCapabilities();

        Assert.Contains(capabilities, item => item.Id == OfficeDocumentReaderBuilderImageExtensions.HandlerId);
        Assert.Contains(capabilities, item => item.Id == OfficeDocumentReaderBuilderNotebookExtensions.HandlerId);
        Assert.Contains(capabilities, item => item.Id == OfficeDocumentReaderBuilderSubtitleExtensions.HandlerId);
    }

    [Fact]
    public void MediaAdapters_DoNotCaptureUnrelatedTextJsonOrUnknownInputsByKind() {
        OfficeDocumentReader allReader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        OfficeDocumentReadResult text = allReader.ReadDocument(Encoding.UTF8.GetBytes("ordinary text"), "notes.txt");
        OfficeDocumentReadResult json = allReader.ReadDocument(Encoding.UTF8.GetBytes("{\"value\":42}"), "payload.bin");
        OfficeDocumentReadResult unknown = allReader.ReadDocument(new byte[] { 0, 1, 2, 3 }, "payload.dat");

        Assert.Contains("ordinary text", text.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeDocumentReaderBuilderSubtitleExtensions.HandlerId, text.CapabilitiesUsed);
        Assert.Contains("officeimo.reader.json", json.CapabilitiesUsed);
        Assert.DoesNotContain(OfficeDocumentReaderBuilderNotebookExtensions.HandlerId, json.CapabilitiesUsed);
        Assert.DoesNotContain(OfficeDocumentReaderBuilderImageExtensions.HandlerId, unknown.CapabilitiesUsed);
    }

    [Fact]
    public void NotebookOnlyRegistration_DoesNotCaptureGenericJsonByDetectedKind() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes("{\"value\":42}"), "payload.bin");

        Assert.DoesNotContain(OfficeDocumentReaderBuilderNotebookExtensions.HandlerId, result.CapabilitiesUsed);
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

    private static byte[] CreateWebpExtendedHeader(int width, int height) {
        var bytes = new byte[30];
        Encoding.ASCII.GetBytes("RIFF").CopyTo(bytes, 0);
        BitConverter.GetBytes(22).CopyTo(bytes, 4);
        Encoding.ASCII.GetBytes("WEBPVP8X").CopyTo(bytes, 8);
        BitConverter.GetBytes(10).CopyTo(bytes, 16);
        WriteUInt24LittleEndian(bytes, 24, width - 1);
        WriteUInt24LittleEndian(bytes, 27, height - 1);
        return bytes;
    }

    private static void WriteUInt24LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
    }
}
