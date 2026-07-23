using OfficeIMO.Drawing;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Notebook;
using OfficeIMO.Reader.Subtitles;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderMediaAdapterTests {
    [Fact]
    public void DocumentResultFactory_PreservesTheFiveParameterAdapterOverload() {
        Type[] parameterTypes = {
            typeof(IEnumerable<ReaderChunk>),
            typeof(ReaderInputKind),
            typeof(OfficeDocumentSource),
            typeof(IEnumerable<string>),
            typeof(IReadOnlyList<OfficeDocumentAsset>)
        };

        Assert.NotNull(typeof(DocumentReaderEngine).GetMethod(
            nameof(DocumentReaderEngine.CreateDocumentResult),
            parameterTypes));
    }

    [Fact]
    public void DependencyFreeMediaAdapters_PopulateTokenEstimates() {
        OfficeDocumentReadResult[] results = {
            new OfficeDocumentReaderBuilder().AddImageHandler().Build()
                .ReadDocument(CreatePng(1, 1), "token.png"),
            new OfficeDocumentReaderBuilder().AddNotebookHandler().Build()
                .ReadDocument(
                    Encoding.UTF8.GetBytes(
                        "{\"cells\":[{\"cell_type\":\"markdown\",\"source\":\"Token estimate\"}]," +
                        "\"metadata\":{},\"nbformat\":4,\"nbformat_minor\":5}"),
                    "token.ipynb"),
            new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build()
                .ReadDocument(
                    Encoding.UTF8.GetBytes("1\n00:00:00,000 --> 00:00:01,000\nToken estimate\n"),
                    "token.srt")
        };

        Assert.All(results, result => Assert.All(result.Chunks, chunk => {
            string projection = chunk.Markdown ?? chunk.Text;
            int expected = projection.Length == 0 ? 0 : Math.Max(1, (projection.Length + 3) / 4);
            Assert.Equal(expected, chunk.TokenEstimate);
        }));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void AdapterProjection_DoesNotSplitSurrogatePairsAtBoundaries(int maxChars) {
        const string value = "a😀b";

        IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(value, maxChars);

        Assert.Equal(value, string.Concat(parts));
        Assert.All(parts, AssertContainsOnlyCompleteSurrogatePairs);
        Assert.Contains("😀", parts);
    }

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
    public void ImageAdapter_BoundsMetadataProjectionAtReaderMaxChars() {
        string sourceName = new string('n', 600) + ".png";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            CreatePng(1, 1),
            sourceName,
            new ReaderOptions { MaxChars = 256 });

        Assert.True(result.Chunks.Count > 1);
        Assert.All(result.Chunks, chunk => {
            Assert.InRange(chunk.Text.Length, 0, 256);
            Assert.InRange(chunk.Markdown?.Length ?? 0, 0, 256);
            Assert.Contains(chunk.Warnings!, warning => warning.Contains("MaxChars", StringComparison.Ordinal));
        });
        Assert.Contains(sourceName, string.Concat(result.Chunks.Select(chunk => chunk.Text)), StringComparison.Ordinal);
        Assert.Contains(
            sourceName,
            string.Concat(result.Chunks.Select(chunk => chunk.Markdown ?? string.Empty)),
            StringComparison.Ordinal);
        Assert.Single(result.Visuals);
        Assert.Single(result.Assets);
    }

    [Fact]
    public void ImageAdapter_RejectsAnImageExtensionWithoutAnImageSignature() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes("not an image"), "renamed.png"));
    }

    [Fact]
    public void ImageAdapter_RejectsPngSignatureWithoutAnIhdrChunk() {
        byte[] validPng = CreatePng(1, 1);
        var malformedPng = new byte[33];
        Array.Copy(validPng, malformedPng, 8);
        Encoding.ASCII.GetBytes("FAKE").CopyTo(malformedPng, 12);
        malformedPng[19] = 3;
        malformedPng[23] = 2;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(malformedPng, "malformed.png"));
    }

    [Fact]
    public void ImageAdapter_RejectsIconWithoutAnInBoundsImagePayload() {
        var malformedIcon = new byte[22];
        malformedIcon[2] = 0x01;
        malformedIcon[4] = 0x01;
        malformedIcon[6] = 16;
        malformedIcon[7] = 16;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(malformedIcon, "malformed.ico"));
    }

    [Fact]
    public void ImageAdapter_RejectsOverflowingEmfDimensionsAsUnsupported() {
        byte[] emf = CreateCompleteEmf(1, 1);
        WriteUInt32LittleEndian(emf, 8, int.MinValue);
        WriteUInt32LittleEndian(emf, 12, int.MinValue);
        WriteUInt32LittleEndian(emf, 16, int.MaxValue);
        WriteUInt32LittleEndian(emf, 20, int.MaxValue);
        WriteUInt32LittleEndian(emf, 24, int.MinValue);
        WriteUInt32LittleEndian(emf, 28, int.MinValue);
        WriteUInt32LittleEndian(emf, 72, int.MaxValue);
        WriteUInt32LittleEndian(emf, 76, int.MaxValue);
        WriteUInt32LittleEndian(emf, 80, 1);
        WriteUInt32LittleEndian(emf, 84, 1);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(emf, "malformed.emf"));
    }

    [Fact]
    public void ImageAdapter_RejectsOverflowingBmpHeightAsUnsupported() {
        var bmp = new byte[54];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteUInt32LittleEndian(bmp, 14, 40);
        WriteUInt32LittleEndian(bmp, 18, 1);
        WriteUInt32LittleEndian(bmp, 22, int.MinValue);
        WriteUInt16LittleEndian(bmp, 26, 1);
        WriteUInt16LittleEndian(bmp, 28, 24);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(bmp, "malformed.bmp"));
    }

    [Fact]
    public void ImageAdapter_RejectsBmpWithTruncatedDeclaredDibHeader() {
        var bmp = new byte[42];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteUInt32LittleEndian(bmp, 14, 40);
        WriteUInt32LittleEndian(bmp, 18, 2);
        WriteUInt32LittleEndian(bmp, 22, 3);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(bmp, "truncated.bmp"));
    }

    [Fact]
    public void ImageAdapter_RejectsBmpWithInvalidPlaneAndBitDepthFields() {
        var bmp = new byte[54];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteUInt32LittleEndian(bmp, 14, 40);
        WriteUInt32LittleEndian(bmp, 18, 2);
        WriteUInt32LittleEndian(bmp, 22, 3);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(bmp, "invalid-fields.bmp"));
    }

    [Fact]
    public void ImageAdapter_RejectsJpegWithoutAFrameHeader() {
        byte[] truncatedJpeg = { 0xFF, 0xD8, 0x00, 0x00 };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(truncatedJpeg, "truncated.jpg"));
    }

    [Fact]
    public void ImageAdapter_RejectsGifWithoutACompleteLogicalScreenDescriptor() {
        var truncatedGif = new byte[10];
        Encoding.ASCII.GetBytes("GIF89a").CopyTo(truncatedGif, 0);
        truncatedGif[6] = 1;
        truncatedGif[8] = 1;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(truncatedGif, "truncated.gif"));
    }

    [Fact]
    public void ImageAdapter_RejectsGifWithATruncatedDeclaredGlobalColorTable() {
        var truncatedGif = new byte[18];
        Encoding.ASCII.GetBytes("GIF89a").CopyTo(truncatedGif, 0);
        truncatedGif[6] = 1;
        truncatedGif[8] = 1;
        truncatedGif[10] = 0x80;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(truncatedGif, "truncated-table.gif"));
    }

    [Fact]
    public void ImageAdapter_IdentifiesWebpWithAnImageChunk() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(CreateWebpExtendedContainer(3, 2), "image.webp");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/webp", asset.MediaType);
        Assert.Equal(3, asset.Width);
        Assert.Equal(2, asset.Height);
    }

    [Fact]
    public void ImageAdapter_RejectsWebpWithoutAnImageChunk() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(CreateWebpExtendedHeaderOnly(3, 2), "header-only.webp"));
    }

    [Fact]
    public void ImageAdapter_RejectsWebpWhenCanvasAndImageDimensionsDiffer() {
        byte[] malformedWebp = CreateWebpExtendedContainer(3, 2);
        WriteUInt24LittleEndian(malformedWebp, 24, 3);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(malformedWebp, "mismatched-canvas.webp"));
    }

    [Fact]
    public void ImageAdapter_RejectsWebpWithAnInvalidContainerLength() {
        byte[] malformedWebp = CreateWebpExtendedContainer(3, 2);
        WriteUInt32LittleEndian(malformedWebp, 4, 0);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(malformedWebp, "malformed.webp"));
    }

    [Fact]
    public void ImageAdapter_RejectsPcxWithInvalidLayoutFields() {
        var pcx = new byte[128];
        pcx[0] = 0x0A;
        pcx[1] = 0x05;
        pcx[2] = 0x01;
        pcx[8] = 1;
        pcx[10] = 1;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(pcx, "malformed.pcx"));
    }

    [Fact]
    public void ImageAdapter_IdentifiesBigTiffFromItsHeader() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(CreateBigTiff(7, 5), "image.tiff");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/tiff", asset.MediaType);
        Assert.Equal(7, asset.Width);
        Assert.Equal(5, asset.Height);
    }

    [Fact]
    public void ImageAdapter_ExportsDefaultDpiForUnitlessBigTiffResolution() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            CreateBigTiff(7, 5, dpi: 300, resolutionUnit: 1),
            "unitless.tiff");

        Assert.Contains(result.Metadata, item => item.Name == "DpiX" && item.Value == "96");
        Assert.Contains(result.Metadata, item => item.Name == "DpiY" && item.Value == "96");
    }

    [Fact]
    public void ImageAdapter_RejectsTruncatedBigTiffDirectory() {
        byte[] tiff = CreateBigTiff(7, 5);
        Array.Resize(ref tiff, 64);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(tiff, "truncated.tiff"));
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
    public void ImageAdapter_IdentifiesContentVerifiedSvgAfterLongCommentPreamble() {
        string svg = "<!--" + new string('x', 5000) +
            "--><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"3\" height=\"2\"/>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(svg), "long-preamble.svg");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/svg+xml", asset.MediaType);
        Assert.Equal(3, asset.Width);
        Assert.Equal(2, asset.Height);
    }

    [Fact]
    public void ImageAdapter_RejectsInvalidSvgDespiteItsExtension() {
        const string svg =
            "<!DOCTYPE svg [<!ENTITY xxe SYSTEM \"file:///invalid\">]>" +
            "<svg xmlns=\"http://www.w3.org/2000/svg\">&xxe;</svg>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes(svg), "invalid.svg"));
    }

    [Fact]
    public void ImageAdapter_IdentifiesContentVerifiedUtf16Svg() {
        const string svg = "\uFEFF<?xml version=\"1.0\" encoding=\"UTF-16\"?><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"4\" height=\"3\"/>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.Unicode.GetBytes(svg), "image.svg");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/svg+xml", asset.MediaType);
        Assert.Equal(4, asset.Width);
        Assert.Equal(3, asset.Height);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageAdapter_IdentifiesContentVerifiedUtf32SvgWithoutBom(bool bigEndian) {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"6\" height=\"5\"/>";
        var encoding = new UTF32Encoding(bigEndian, byteOrderMark: false, throwOnInvalidCharacters: true);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(encoding.GetBytes(svg), "image.svg");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/svg+xml", asset.MediaType);
        Assert.Equal(6, asset.Width);
        Assert.Equal(5, asset.Height);
    }

    [Fact]
    public void ImageAdapter_RejectsMalformedSvgAfterAValidRootStartTag() {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"5\" height=\"4\"><unclosed";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes(svg), "header.svg"));
    }

    [Fact]
    public void ImageAdapter_LeavesOversizedSvgDimensionsUnknown() {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"3000000000\" height=\"4\"/>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(svg), "oversized.svg");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Null(asset.Width);
        Assert.Equal(4, asset.Height);
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
    public void NotebookAdapter_AcceptsUtf8Bom() {
        const string notebook = "{\"cells\":[{\"cell_type\":\"markdown\",\"source\":\"BOM notebook\"}]," +
            "\"metadata\":{},\"nbformat\":4,\"nbformat_minor\":5}";
        byte[] content = Encoding.UTF8.GetBytes(notebook);
        var bytes = new byte[content.Length + 3];
        bytes[0] = 0xEF;
        bytes[1] = 0xBB;
        bytes[2] = 0xBF;
        Buffer.BlockCopy(content, 0, bytes, 3, content.Length);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "bom.ipynb");

        Assert.Equal("BOM notebook", Assert.Single(result.Chunks).Text);
        Assert.Contains(result.Metadata, item => item.Name == "NbFormat" && item.Value == "4");
    }

    [Fact]
    public void NotebookAdapter_BoundsFenceLanguageMetadataBeforeReusingIt() {
        string oversizedLanguage = new string('p', 1000);
        string notebook = "{\"cells\":[{\"cell_type\":\"code\",\"source\":\"one\"}," +
            "{\"cell_type\":\"code\",\"source\":\"two\"}],\"metadata\":{\"kernelspec\":{\"language\":\"" +
            oversizedLanguage + "\"}},\"nbformat\":4,\"nbformat_minor\":5}";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "bounded-language.ipynb");

        string expectedLanguage = new string('p', 64);
        Assert.Equal(expectedLanguage, Assert.Single(result.Metadata, item => item.Name == "Language").Value);
        Assert.All(result.Chunks, chunk => {
            Assert.Contains("```" + expectedLanguage, chunk.Markdown, StringComparison.Ordinal);
            Assert.DoesNotContain(new string('p', 65), chunk.Markdown, StringComparison.Ordinal);
        });
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
        Assert.StartsWith("    indented code", chunk.Text, StringComparison.Ordinal);
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
    public void NotebookAdapter_TruncatesSourceAndOutputAtCompleteUnicodeScalars() {
        const string notebook = """
            {
              "cells": [
                { "cell_type": "markdown", "source": "😀source" },
                {
                  "cell_type": "code",
                  "source": "",
                  "outputs": [ { "output_type": "stream", "text": "😀output" } ]
                }
              ],
              "metadata": {},
              "nbformat": 4,
              "nbformat_minor": 0
            }
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddNotebookHandler(new ReaderNotebookOptions {
                MaxCellCharacters = 1,
                MaxOutputCharactersPerCell = 1
            })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "unicode.ipynb");

        Assert.Equal(2, result.Chunks.Count);
        Assert.All(result.Chunks, chunk => {
            AssertContainsOnlyCompleteSurrogatePairs(chunk.Text);
            AssertContainsOnlyCompleteSurrogatePairs(chunk.Markdown ?? string.Empty);
            Assert.Contains("😀", chunk.Text, StringComparison.Ordinal);
            Assert.DoesNotContain("source", chunk.Text, StringComparison.Ordinal);
            Assert.DoesNotContain("output", chunk.Text, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void NotebookAdapter_FallsBackToPlainTextWhenMarkdownOutputIsEmpty() {
        const string notebook = """
            {
              "cells": [
                {
                  "cell_type": "code",
                  "source": "display(value)",
                  "outputs": [
                    {
                      "output_type": "display_data",
                      "data": { "text/markdown": [], "text/plain": ["plain fallback"] }
                    }
                  ]
                }
              ],
              "metadata": {},
              "nbformat": 4,
              "nbformat_minor": 5
            }
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(notebook), "fallback.ipynb");

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Contains("plain fallback", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("plain fallback", chunk.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void NotebookAdapter_SplitsCellProjectionAtReaderMaxChars() {
        string source = new string('a', 600);
        string notebook = "{\"cells\":[{\"cell_type\":\"code\",\"source\":\"" + source +
            "\",\"outputs\":[]}],\"metadata\":{\"kernelspec\":{\"language\":\"python\"}}," +
            "\"nbformat\":4,\"nbformat_minor\":5}";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddNotebookHandler().Build();
        byte[] bytes = Encoding.UTF8.GetBytes(notebook);

        OfficeDocumentReadResult unsplit = reader.ReadDocument(bytes, "split.ipynb");

        OfficeDocumentReadResult result = reader.ReadDocument(
            bytes,
            "split.ipynb",
            new ReaderOptions { MaxChars = 256 });

        Assert.Equal(3, result.Chunks.Count);
        Assert.All(result.Chunks, chunk => {
            Assert.InRange(chunk.Text.Length, 1, 256);
            Assert.InRange(chunk.Markdown.Length, 1, 256);
            Assert.Equal(0, chunk.Location.SourceBlockIndex);
            Assert.Contains(chunk.Warnings!, warning => warning.Contains("MaxChars", StringComparison.Ordinal));
        });
        Assert.Equal(source, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.Equal(unsplit.Markdown, string.Concat(result.Chunks.Select(chunk => chunk.Markdown)));
        Assert.Equal(unsplit.Markdown, result.Markdown);
    }

    [Fact]
    public void NotebookAdapter_PreservesContinuationWhenProcessorReplacesSplitChunks() {
        string source = new string('a', 600);
        string notebook = "{\"cells\":[{\"cell_type\":\"markdown\",\"source\":\"" + source +
            "\"}],\"metadata\":{},\"nbformat\":4,\"nbformat_minor\":5}";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddNotebookHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("replace-split-chunks", (document, _) => {
                document.Chunks = document.Chunks.Select(chunk => new ReaderChunk {
                    Id = chunk.Id,
                    Kind = chunk.Kind,
                    Location = chunk.Location,
                    Text = chunk.Text.Replace('a', 'b'),
                    Markdown = chunk.Markdown
                }).ToArray();
                return document;
            }))
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(notebook),
            "processed-split.ipynb",
            new ReaderOptions { MaxChars = 256 });

        Assert.Equal(new string('b', 600), result.Markdown);
    }

    [Fact]
    public void NotebookAdapter_PreservesTruncationDiagnosticWhenRetainedCellPrefixIsWhitespace() {
        string source = new string(' ', 300) + "omitted content";
        string notebook = "{\"cells\":[{\"cell_type\":\"markdown\",\"source\":\"" + source +
            "\"}],\"metadata\":{},\"nbformat\":4,\"nbformat_minor\":5}";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddNotebookHandler(new ReaderNotebookOptions { MaxCellCharacters = 256 })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(notebook),
            "truncated.ipynb");

        Assert.Empty(result.Chunks);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(
            result.Diagnostics,
            item => item.Code == "notebook-content-truncated");
        Assert.Contains("MaxCellCharacters", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal(0, diagnostic.Location?.SourceBlockIndex);
        Assert.Equal("256", diagnostic.Attributes["maxCellCharacters"]);
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

    [Theory]
    [InlineData("1 < 2", "1 < 2")]
    [InlineData("1 < 2 > 0", "1 < 2 > 0")]
    [InlineData("Keep <unknown> literal", "Keep <unknown> literal")]
    [InlineData("<i>Hello</i> <v Speaker>there</v>", "Hello there")]
    public void SubtitleAdapter_StripsCueTagsWithoutDroppingLiteralComparisons(string cueText, string expected) {
        string srt = "1\n00:00:00,000 --> 00:00:01,000\n" + cueText + "\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "literal.srt");

        Assert.Equal(expected, Assert.Single(result.Chunks).Text);
    }

    [Fact]
    public void SubtitleAdapter_SplitsCueProjectionAtReaderMaxChars() {
        string cueText = new string('x', 1_000);
        string srt = "1\n00:00:00,000 --> 00:00:01,000\n" + cueText + "\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(srt),
            "bounded.srt",
            new ReaderOptions { MaxChars = 256 });

        Assert.True(result.Chunks.Count > 1);
        Assert.Equal(cueText, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.All(result.Chunks, chunk => {
            Assert.True(chunk.Text.Length <= 256);
            Assert.True((chunk.Markdown?.Length ?? 0) <= 256);
            Assert.Equal(0, chunk.Location.SourceBlockIndex);
            Assert.Contains(chunk.Warnings!, warning => warning.Contains("MaxChars", StringComparison.Ordinal));
        });
        Assert.StartsWith("**00:00:00.000 → 00:00:01.000**", result.Chunks[0].Markdown, StringComparison.Ordinal);
        Assert.Equal(string.Concat(result.Chunks.Select(chunk => chunk.Markdown)), result.Markdown);
        Assert.Equal(result.Chunks.Count, result.Chunks.Select(chunk => chunk.Id).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void SubtitleAdapter_SplitsEscapedMarkdownAtMatchingSourceRanges() {
        string cueText = string.Concat(Enumerable.Repeat("123456<789&ABC>DEF", 40));
        string srt = "1\n00:00:00,000 --> 00:00:01,000\n" + cueText + "\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddSubtitleHandler(new ReaderSubtitleOptions { IncludeTimestampsInMarkdown = false })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(srt),
            "aligned.srt",
            new ReaderOptions { MaxChars = 256 });

        Assert.True(result.Chunks.Count > 1);
        Assert.Equal(cueText, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.Equal(EscapeSubtitleMarkdown(cueText),
            string.Concat(result.Chunks.Select(chunk => chunk.Markdown)));
        Assert.All(result.Chunks, chunk =>
            Assert.Equal(EscapeSubtitleMarkdown(chunk.Text), chunk.Markdown));
    }

    [Fact]
    public void AdapterSnapshot_ReappliesTheRegisteredDefaultInputLimit() {
        const string extension = ".adapterlimit";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.adapter-limit",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                DefaultMaxInputBytes = 16,
                ReadDocumentStream = (_input, sourceName, options, cancellationToken) => {
                    using var oversized = new MemoryStream(new byte[17], writable: false);
                    _ = DocumentReaderEngine.ReadAdapterInput(oversized, sourceName, options, cancellationToken);
                    return new OfficeDocumentReadResult { Kind = ReaderInputKind.Text };
                }
            })
            .Build();

        IOException exception = Assert.Throws<IOException>(() => reader.ReadDocument(
            new byte[1],
            "sample" + extension));

        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.Ordinal);
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
    public void SubtitleAdapter_DoesNotTreatAnSrtIdentifierStartingWithWebVttAsAHeader() {
        const string srt = "WEBVTT_note\n00:00:00,000 --> 00:00:01,000\nFirst cue\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "captions.srt");

        Assert.Equal("First cue", Assert.Single(result.Chunks).Text);
        Assert.Contains(result.Metadata, item => item.Name == "Format" && item.Value == "srt");
    }

    [Fact]
    public void SubtitleAdapter_DoesNotTreatAnSrtIdentifierWithWebVttPrefixAsAHeader() {
        const string srt = "WEBVTT cue\n00:00:00,000 --> 00:00:01,000\nFirst cue\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "captions.srt");

        Assert.Equal("First cue", Assert.Single(result.Chunks).Text);
        Assert.Contains(result.Metadata, item => item.Name == "Format" && item.Value == "srt");
    }

    [Fact]
    public void SubtitleAdapter_TruncatesCueAtACompleteUnicodeScalar() {
        const string srt = "1\n00:00:00,000 --> 00:00:01,000\n😀suffix\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddSubtitleHandler(new ReaderSubtitleOptions { MaxCueCharacters = 1 })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "unicode.srt");

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal("😀", chunk.Text);
        AssertContainsOnlyCompleteSurrogatePairs(chunk.Markdown ?? string.Empty);
        Assert.Contains(chunk.Warnings!, warning => warning.Contains("MaxCueCharacters", StringComparison.Ordinal));
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
    public void SubtitleAdapter_DiagnosesADanglingCueIdentifier() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes("42"), "truncated.srt");

        Assert.Empty(result.Chunks);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("subtitle-invalid-block", diagnostic.Code);
        Assert.Contains("invalid timing line at line 1", diagnostic.Message, StringComparison.Ordinal);
        Assert.True(diagnostic.IsRecoverable);
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

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes("{\"value\":42}"), "payload.bin"));
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

    private static void AssertContainsOnlyCompleteSurrogatePairs(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index])) {
                Assert.True(index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]));
                index++;
            } else {
                Assert.False(char.IsLowSurrogate(value[index]));
            }
        }
    }

    private static byte[] CreateWebpExtendedContainer(int width, int height) {
        return OfficeWebpCodec.Encode(
            new OfficeRasterImage(width, height, OfficeColor.FromRgb(32, 64, 128)),
            dpiX: 96D,
            dpiY: 96D);
    }

    private static byte[] CreateWebpExtendedHeaderOnly(int width, int height) {
        var bytes = new byte[30];
        Encoding.ASCII.GetBytes("RIFF").CopyTo(bytes, 0);
        WriteUInt32LittleEndian(bytes, 4, 22);
        Encoding.ASCII.GetBytes("WEBPVP8X").CopyTo(bytes, 8);
        WriteUInt32LittleEndian(bytes, 16, 10);
        WriteUInt24LittleEndian(bytes, 24, width - 1);
        WriteUInt24LittleEndian(bytes, 27, height - 1);
        return bytes;
    }

    private static byte[] CreateBigTiff(
        int width,
        int height,
        int? dpi = null,
        int resolutionUnit = 2) {
        int entryCount = dpi.HasValue ? 5 : 2;
        var bytes = new byte[32 + (entryCount * 20)];
        bytes[0] = (byte)'I';
        bytes[1] = (byte)'I';
        WriteUInt16LittleEndian(bytes, 2, 43);
        WriteUInt16LittleEndian(bytes, 4, 8);
        WriteUInt64LittleEndian(bytes, 8, 16);
        WriteUInt64LittleEndian(bytes, 16, (ulong)entryCount);
        WriteBigTiffLongEntry(bytes, 24, 256, width);
        WriteBigTiffLongEntry(bytes, 44, 257, height);
        if (dpi.HasValue) {
            WriteBigTiffRationalEntry(bytes, 64, 282, dpi.Value, 1);
            WriteBigTiffRationalEntry(bytes, 84, 283, dpi.Value, 1);
            WriteBigTiffShortEntry(bytes, 104, 296, resolutionUnit);
        }
        return bytes;
    }

    private static void WriteBigTiffLongEntry(byte[] bytes, int offset, int tag, int value) {
        WriteUInt16LittleEndian(bytes, offset, tag);
        WriteUInt16LittleEndian(bytes, offset + 2, 4);
        WriteUInt64LittleEndian(bytes, offset + 4, 1);
        WriteUInt32LittleEndian(bytes, offset + 12, value);
    }

    private static void WriteBigTiffRationalEntry(
        byte[] bytes,
        int offset,
        int tag,
        int numerator,
        int denominator) {
        WriteUInt16LittleEndian(bytes, offset, tag);
        WriteUInt16LittleEndian(bytes, offset + 2, 5);
        WriteUInt64LittleEndian(bytes, offset + 4, 1);
        WriteUInt32LittleEndian(bytes, offset + 12, numerator);
        WriteUInt32LittleEndian(bytes, offset + 16, denominator);
    }

    private static void WriteBigTiffShortEntry(byte[] bytes, int offset, int tag, int value) {
        WriteUInt16LittleEndian(bytes, offset, tag);
        WriteUInt16LittleEndian(bytes, offset + 2, 3);
        WriteUInt64LittleEndian(bytes, offset + 4, 1);
        WriteUInt16LittleEndian(bytes, offset + 12, value);
    }

    private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32LittleEndian(byte[] bytes, int offset, int value) {
        for (int index = 0; index < 4; index++) {
            bytes[offset + index] = (byte)(value >> (index * 8));
        }
    }

    private static void WriteUInt64LittleEndian(byte[] bytes, int offset, ulong value) {
        for (int index = 0; index < 8; index++) {
            bytes[offset + index] = (byte)(value >> (index * 8));
        }
    }

    private static void WriteUInt24LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
    }
}
