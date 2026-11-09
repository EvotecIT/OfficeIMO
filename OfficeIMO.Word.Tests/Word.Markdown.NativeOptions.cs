using System.Net;
using System.Net.Sockets;
using System.Text;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Internal;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class MarkdownNativeOptionsTests {
    [Fact]
    public async Task NativeRemoteImageOptInDownloadsOnceAndLeavesCallerOptionsReusable() {
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        byte[] png = await OfficeRuntimeCompatibility.ReadAllBytesAsync(Path.Combine(AppContext.BaseDirectory, "NativeOptionAssets", "OfficeIMO.png"), deadline.Token);
        Task server = ServeImageAsync(listener, png, deadline.Token);
        try {
            string uri = $"http://127.0.0.1:{port}/image.png";
            using MarkdownDocument source = MarkdownDocument.Parse($"![first]({uri})\n\n![second]({uri})");
            var options = new MarkdownWordExportOptions { AllowRemoteImages = true, MaximumRemoteImageBytes = png.Length,
                RemoteImageLoadOptions = new OfficeRemoteImageLoadOptions { AllowPrivateNetworkAddresses = true } };
            OfficeConversionResult<WordDocument> converted = await source.ToWordAsync(options, deadline.Token);
            using WordDocument word = converted.Value;
            await server;
            Assert.True(word.Images.Count == 2, string.Join("\n", converted.Report.Diagnostics.Select(finding => finding.Message)));
            Assert.Null(options.RemoteImageResolver);
            Assert.True(source.Blocks.Count > 0);
        } finally { listener.Stop(); try { await server; } catch (OperationCanceledException) { } catch (SocketException) { } }
    }

    [Fact]
    public async Task RemoteImageByteRejectionRetainsLossAndDoesNotReplaceTheDestination() {
        using MarkdownDocument source = MarkdownDocument.Parse("![image](https://example.invalid/image.png)");
        using var output = new MemoryStream();
        var options = new MarkdownWordExportOptions { AllowRemoteImages = true, MaximumRemoteImageBytes = 0, FailOnLoss = true };
        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() => source.SaveAsWordAsync(output, options));
        Assert.True(error.Report.HasLoss);
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task NativeThemeProjectsRolesBeforeAuthoredMarkdownAndKeepsTemplatePrecedence() {
        OfficeTheme native = OfficeTheme.Modern.WithFonts(OfficeTheme.Modern.Fonts.With(
            heading: OfficeTheme.Modern.Fonts.Heading.With(family: "Georgia"), body: OfficeTheme.Modern.Fonts.Body.With(family: "Courier New")))
            .WithTypeScale(OfficeTheme.Modern.TypeScale.With(heading1: OfficeLength.Points(29), body: OfficeLength.Points(14)))
            .WithTextStyle(OfficeTheme.Modern.TextStyle.With(headingColor: OfficeColor.Parse("#892345"),
                linkColor: OfficeColor.Parse("#136789"), headingBold: false))
            .WithTableStyle(OfficeTheme.Modern.TableStyle.With(headerBold: false, headerFill: OfficeColor.Parse("#CCE8F5")))
            .WithPage(OfficeTheme.Modern.Page.With(marginLeft: OfficeLength.Points(37)));
        MarkdownVisualTheme mapped = MarkdownVisualTheme.FromOfficeTheme(native);
        using MarkdownDocument source = MarkdownDocument.Parse("# Heading **bold** [link](https://example.test)\n\nBody *italic*.\n\n| Header | Count |\n| --- | --- |\n| Tablecell | 2 |\n");
        var converted = await source.ToWordAsync(new MarkdownWordExportOptions { Theme = mapped });
        using WordDocument created = converted.Value;
        using WordDocument word = WordDocument.Load(new MemoryStream(created.ToBytes()));
        var styles = word._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var headingStyle = styles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Heading1");
        var normalStyle = styles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Normal");
        Assert.Equal("58", headingStyle.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("892345", headingStyle.StyleRunProperties.Color!.Val!.Value);
        Assert.False(headingStyle.StyleRunProperties.Bold?.Val?.Value ?? false);
        Assert.Equal("28", normalStyle.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal(OfficeLength.Points(37), word.PageSetup.Margins.Left);
        var fonts = word._wordprocessingDocument.MainDocumentPart.ThemePart!.Theme.ThemeElements!.FontScheme!;
        Assert.Equal("Georgia", fonts.MajorFont!.LatinFont!.Typeface!.Value);
        Assert.Equal("Courier New", fonts.MinorFont!.LatinFont!.Typeface!.Value);
        var body = word.OpenXmlDocument.MainDocumentPart.Document.Body!;
        var heading = Assert.Single(body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>(),
            paragraph => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Heading1");
        var bold = Assert.Single(heading.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>(), run => run.InnerText == "bold");
        Assert.True(bold.RunProperties?.Bold != null && (bold.RunProperties.Bold.Val?.Value ?? true));
        Assert.Null(heading.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().First(run => run.InnerText.Contains("Heading")).RunProperties?.Color);
        Assert.Contains(word.HyperLinks, link => link.Uri?.ToString() == "https://example.test/");
        var italic = Assert.Single(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>(), run => run.InnerText == "italic");
        Assert.True(italic.RunProperties?.Italic != null && (italic.RunProperties.Italic.Val?.Value ?? true));
        Assert.Empty(word.ValidateDocument());
        string? HeaderFill(WordDocument document) => WordRunFormatting.ResolveShadingFillColorHex(
            WordRunFormatting.ResolveTableCellShading(document.Tables.Single().Rows[0].Cells[0]), document)?.ToUpperInvariant();
        Assert.Equal("CCE8F5", HeaderFill(word));

        MarkdownVisualTheme customized = mapped.WithColors(heading: "#004400", tableHeaderBackground: "#0000FF")
            .WithTable(table => { table.CellPaddingY = OfficeLength.Points(8); table.UseRowStripes = false; });
        customized.Table = customized.Table.Clone();
        customized.Table.CellPaddingX = OfficeLength.Points(13);
        using WordDocument edited = source.ToWord(new MarkdownWordExportOptions { Theme = customized });
        using WordDocument editedArtifact = WordDocument.Load(new MemoryStream(edited.ToBytes()));
        var editedHeading = editedArtifact.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Heading1");
        Assert.Equal("004400", editedHeading.StyleRunProperties!.Color!.Val!.Value);
        Assert.Equal("58", editedHeading.StyleRunProperties.FontSize!.Val!.Value);
        Assert.Equal("0000FF", HeaderFill(editedArtifact));
        Assert.Equal(OfficeLength.Points(13), editedArtifact.Tables.Single().Rows[1].Cells[0].MarginLeft);
        Assert.Equal(OfficeLength.Points(37), editedArtifact.PageSetup.Margins.Left);
        Assert.Equal("136789", editedArtifact.HyperLinks.Single().Runs.Single().Color?.ToRgbHex());

        using MarkdownDocument tableRoles = MarkdownDocument.Parse("| Header **Emphasis** `Code` [Link](https://example.test/table) " +
            "**![Fallback](not-supplied-theme-cell-image.png)** **[^cell]** | Other |\n" +
            "| --- | --- |\n| Body *detail* | Data |\n\n| Header only |\n| --- |\n\nFollowing body.\n\n[^cell]: Cell note.\n");
        var definitions = new DefinitionListBlock();
        definitions.AddEntry("Term", "Definition");
        tableRoles.Blocks.OfType<TableBlock>().First().GetHeaderCell(0)!.ChildBlocks.Add(definitions);
        foreach (var sample in new[] {
            (Theme: OfficeTheme.Boardroom, Fill: "14335C", Foreground: "FFFFFF", Body: "16202E"),
            (Theme: OfficeTheme.Dark, Fill: "1E2A39", Foreground: "E7ECF2", Body: "0F1620")
        }) {
            using WordDocument themed = tableRoles.ToWord(new MarkdownWordExportOptions { Theme = MarkdownVisualTheme.FromOfficeTheme(sample.Theme) });
            using WordDocument artifact = WordDocument.Load(new MemoryStream(themed.ToBytes()));
            Assert.Empty(artifact.ValidateDocument());
            WordTable table = artifact.Tables.First();
            WordTableCell header = table.Rows[0].Cells[0];
            var headerRun = Assert.Single(header._tableCell.Descendants<W.Run>(), run => run.InnerText.Trim() == "Header");
            Assert.Equal((sample.Fill, sample.Foreground),
                (WordRunFormatting.ResolveShadingFillColorHex(WordRunFormatting.ResolveTableCellShading(header), artifact)?.ToUpperInvariant(),
                    headerRun.RunProperties?.Color?.Val?.Value));
            Assert.Equal("https://example.test/table", artifact.HyperLinks.Single().Uri?.ToString());
            Assert.Equal(sample.Theme.TextStyle.LinkColor.Resolve(sample.Theme).ToRgbHex(), artifact.HyperLinks.Single().Runs.Single().Color?.ToRgbHex());
            Assert.True(Assert.Single(header._tableCell.Descendants<W.Run>(), run => run.InnerText == "Emphasis").RunProperties!.Bold!.Val?.Value ?? true);
            var code = Assert.Single(header._tableCell.Descendants<W.Run>(), run => run.InnerText == "Code");
            Assert.Equal(WordFontResolver.Resolve("monospace") ?? "Consolas", code.RunProperties!.RunFonts!.Ascii!.Value);
            Assert.Null(code.RunProperties!.RunFonts!.AsciiTheme);
            Assert.Equal(sample.Foreground, code.RunProperties?.Color?.Val?.Value);
            var fallback = Assert.Single(header._tableCell.Descendants<W.Run>(), run => run.InnerText == "Fallback");
            Assert.Equal(sample.Foreground, fallback.RunProperties!.Color!.Val!.Value);
            Assert.True(fallback.RunProperties.Bold!.Val?.Value ?? true);
            var reference = Assert.Single(header._tableCell.Descendants<W.Run>(), run => run.GetFirstChild<W.FootnoteReference>() != null);
            Assert.Equal(sample.Foreground, reference.RunProperties!.Color!.Val!.Value);
            Assert.True(reference.RunProperties.Bold!.Val?.Value ?? true);
            Assert.Equal("FootnoteReference", reference.RunProperties.RunStyle!.Val!.Value);
            var referenceStyle = artifact.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<W.Style>()
                .Single(style => style.StyleId?.Value == "FootnoteReference");
            Assert.Equal(W.VerticalPositionValues.Superscript, referenceStyle.StyleRunProperties!.VerticalTextAlignment!.Val!.Value);
            var separator = Assert.Single(header._tableCell.Descendants<W.Run>(), run => run.InnerText == ": ");
            Assert.Equal(sample.Foreground, separator.RunProperties!.Color!.Val!.Value);
            foreach (W.Paragraph paragraph in table._table.Descendants<W.Paragraph>()) {
                var spacing = paragraph.ParagraphProperties!.SpacingBetweenLines!;
                Assert.Equal(("0", "0", "240"), (spacing.Before!.Value, spacing.After!.Value, spacing.Line!.Value));
            }
            Assert.Equal(sample.Body, table.Rows[1].Cells[0]._tableCell.Descendants<W.Run>().First(run => !string.IsNullOrWhiteSpace(run.InnerText)).RunProperties?.Color?.Val?.Value);
            var following = Assert.Single(artifact.OpenXmlDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>(), paragraph => paragraph.InnerText == "Following body.");
            Assert.All(following.Descendants<W.Run>(), run => Assert.Null(run.RunProperties?.Color));
            var followingNormal = artifact.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
                .Single(style => style.StyleId?.Value == "Normal");
            Assert.Equal(sample.Body, followingNormal.StyleRunProperties!.Color!.Val!.Value);
        }

        using WordDocument explicitFont = source.ToWord(new MarkdownWordExportOptions { Theme = mapped, FontFamily = "Verdana" });
        Assert.Equal("Verdana", explicitFont.Paragraphs.SelectMany(paragraph => paragraph.GetRuns()).First(run => run.Text.Contains("Heading")).FontFamily);
        Assert.All(explicitFont.Tables.Single()._table.Descendants<W.Run>()
            .Where(run => !string.IsNullOrWhiteSpace(run.InnerText)), run => {
                Assert.Equal("Verdana", run.RunProperties!.RunFonts!.Ascii!.Value);
                Assert.Null(run.RunProperties.RunFonts.AsciiTheme);
            });
        using WordDocument plain = source.ToWord();
        var plainHeading = plain._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Heading1");
        Assert.Equal(WordThemeDefaults.PageText(OfficeTheme.Default).ToRgbHex(), plainHeading.StyleRunProperties!.Color!.Val!.Value);
        Assert.True(plainHeading.StyleRunProperties.Bold?.Val?.Value ?? true);

        using WordDocument template = WordDocument.Create();
        template.PageSetup.Margins.Left = OfficeLength.Points(19);
        template.AddParagraph("Keep template");
        var templateNormal = template.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
            .Single(style => style.StyleId?.Value == "Normal");
        templateNormal.StyleRunProperties = new DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties(
            new DocumentFormat.OpenXml.Wordprocessing.RunFonts { Ascii = "Tahoma" },
            new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "AA0011" });
        byte[] templateBytes = template.ToBytes();
        using WordDocument templated = source.ToWord(new MarkdownWordExportOptions { Theme = mapped, Template = template });
        Assert.Equal(OfficeLength.Points(19), templated.PageSetup.Margins.Left);
        var templateHeading = templated._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Heading1");
        Assert.Equal(OfficeTheme.Default.TextStyle.HeadingColor.Resolve(OfficeTheme.Default).ToRgbHex(), templateHeading.StyleRunProperties!.Color!.Val!.Value);
        Assert.Contains(templated.Paragraphs, paragraph => paragraph.Text == "Keep template");
        Assert.All(templated.Tables.Single()._table.Descendants<W.Run>(),
            run => Assert.Null(run.RunProperties?.Color));
        var templatedNormal = templated.OpenXmlDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
            .Single(style => style.StyleId?.Value == "Normal");
        Assert.Equal(("Tahoma", "AA0011"),
            (templatedNormal.StyleRunProperties!.RunFonts!.Ascii!.Value, templatedNormal.StyleRunProperties.Color!.Val!.Value));
        string path = Path.Combine(Path.GetTempPath(), "officeimo-theme-template-" + Guid.NewGuid().ToString("N") + ".docx");
        try {
            File.WriteAllBytes(path, templateBytes);
            using WordDocument fromPath = source.ToWord(new MarkdownWordExportOptions { Theme = customized, TemplatePath = path });
            var pathHeading = fromPath.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Single(style => style.StyleId?.Value == "Heading1");
            Assert.Equal(templateHeading.StyleRunProperties.Color!.Val!.Value, pathHeading.StyleRunProperties!.Color!.Val!.Value);
            Assert.Equal(OfficeLength.Points(19), fromPath.PageSetup.Margins.Left);
            Assert.Equal(templateBytes, File.ReadAllBytes(path));
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    private static async Task ServeImageAsync(TcpListener listener, byte[] image, CancellationToken token) {
        try {
            using CancellationTokenRegistration stopListening = token.Register(listener.Stop);
            using TcpClient client = await listener.AcceptTcpClientAsync();
            using CancellationTokenRegistration closeClient = token.Register(client.Close);
            using NetworkStream stream = client.GetStream();
            var bytes = new List<byte>();
            var buffer = new byte[1024];
            while (!Encoding.ASCII.GetString(bytes.ToArray()).Contains("\r\n\r\n", StringComparison.Ordinal)) {
                int read = await stream.ReadAsync(buffer, 0, buffer.Length, token);
                if (read == 0) throw new EndOfStreamException();
                bytes.AddRange(buffer.Take(read));
                if (bytes.Count > 16384) throw new InvalidDataException("Unexpected loopback request size.");
            }
            byte[] header = Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: {image.Length}\r\nConnection: close\r\n\r\n");
            await stream.WriteAsync(header, 0, header.Length, token);
            await stream.WriteAsync(image, 0, image.Length, token);
        } catch (Exception error) when (token.IsCancellationRequested
            && (error is SocketException || error is IOException || error is ObjectDisposedException)) {
            throw new OperationCanceledException("The loopback image server was canceled.", error, token);
        }
    }
}
