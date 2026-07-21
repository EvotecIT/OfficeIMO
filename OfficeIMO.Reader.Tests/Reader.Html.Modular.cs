using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Markdown.Html;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderHtmlModularTests {
    [Fact]
    public void DocumentReaderHtml_RichDispatch_MapsMetadataStructureLinksFormsAndImages() {
        const string html = "<html><head><title>Rich HTML</title><meta name=\"author\" content=\"OfficeIMO\"/></head><body>"
            + "<h2>Inventory</h2><p>See <a href=\"https://example.test/inventory\">inventory</a>.</p>"
            + "<table><tr><th>Name</th><th>Qty</th></tr><tr><td>Bandage</td><td>4</td></tr></table>"
            + "<form><input type=\"text\" name=\"patient\" value=\"Ada\" required=\"required\"/></form>"
            + "<img alt=\"Tiny image\" src=\"data:image/png;base64,iVBORw0KGgo=\"/></body></html>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "rich.html");

        Assert.Equal("Rich HTML", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        Assert.Contains(result.Blocks, block => block.Kind == "heading" && block.Level == 2 && block.Text == "Inventory");
        ReaderTable table = Assert.Single(result.Tables);
        Assert.Equal("Bandage", table.Rows[0][0]);
        Assert.Equal("https://example.test/inventory", Assert.Single(result.Links).Uri);
        OfficeDocumentFormField form = Assert.Single(result.Forms);
        Assert.Equal("patient", form.Name);
        Assert.True(form.IsRequired);
        OfficeDocumentAsset image = Assert.Single(result.Assets);
        Assert.Equal("image/png", image.MediaType);
        Assert.NotNull(image.PayloadBytes);
        ReaderVisual visual = Assert.Single(result.Visuals);
        Assert.Equal("image", visual.Kind);
        Assert.Equal(image.PayloadHash, visual.PayloadHash);
        stream.Position = 0;
        OfficeDocumentReadResult jsonResult = OfficeDocumentReadResultJson.Deserialize(
            HtmlReaderAdapter.ReadDocumentJson(stream, "rich.html"));
        Assert.Equal(ReaderInputKind.Html, jsonResult.Kind);
        Assert.Contains("officeimo.reader.html.rich-v5", result.CapabilitiesUsed);
    }

    [Fact]
    public void DocumentReaderHtml_RichDispatch_AppliesConfiguredUrlPolicyToLinks() {
        const string html = "<p><a href=\"javascript:alert(1)\">Unsafe</a> <a href=\"https://example.test/safe\">Safe</a></p>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "links.html");

        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("https://example.test/safe", link.Uri);
        Assert.DoesNotContain(result.Links, item => item.Uri?.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase) == true);
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_ResolvesDocumentBaseUriWithoutFilters() {
        const string html = "<html><head><base href=\"https://example.test/docs/\"></head>"
            + "<body><a href=\"guide.html\">Guide</a></body></html>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "guide.html");

        Assert.Equal("https://example.test/docs/guide.html", Assert.Single(result.Links).Uri);
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_ProjectsNestedMediaSourcesOnce() {
        const string html = "<audio><source src=\"chapter.mp3\" type=\"audio/mpeg\"/>Audio fallback</audio>"
            + "<video><source src=\"clip.webm\" type=\"video/webm\"/>"
            + "<source src=\"clip.mp4\" type=\"video/mp4\"/>Video fallback</video>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "media.html");

        Assert.Equal(2, result.Visuals.Count);
        ReaderVisual audio = Assert.Single(result.Visuals, visual => visual.Language == "audio");
        Assert.Equal("chapter.mp3", audio.SourceName);
        Assert.Equal("audio/mpeg", audio.MimeType);
        ReaderVisual video = Assert.Single(result.Visuals, visual => visual.Language == "video");
        Assert.Equal("clip.webm", video.SourceName);
        Assert.Equal("video/webm", video.MimeType);
        Assert.DoesNotContain(result.Visuals, visual => visual.Language == "source");
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_RejectsOversizedInputBeforeLogicalProjection() {
        const string html = "<p>This input exceeds its configured bound.</p>";

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            HtmlReaderAdapter.ReadContentDocument(
                html,
                "bounded-rich.html",
                htmlOptions: ReaderHtmlOptions.CreateUntrustedHtmlProfile(12)));

        Assert.Contains("MaxInputCharacters", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(HtmlBase64ImageHandling.Skip)]
    [InlineData(HtmlBase64ImageHandling.SaveToFile)]
    public void DocumentReaderHtml_RichDispatch_RespectsBase64ImageHandling(HtmlBase64ImageHandling handling) {
        const string html = "<img alt=\"Inline\" src=\"data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==\"/>";
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-html-images-" + Guid.NewGuid().ToString("N"));
        try {
            var options = new ReaderHtmlOptions {
                HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                    Base64Images = handling,
                    Base64ImageOutputDirectory = directory
                }
            };

            OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(
                html,
                "inline-image.html",
                htmlOptions: options);

            Assert.Empty(result.Assets);
            Assert.Empty(result.Visuals);
            if (handling == HtmlBase64ImageHandling.SaveToFile) Assert.Single(Directory.EnumerateFiles(directory));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void DocumentReaderHtml_RichTables_DoNotFoldNestedRowsIntoParentTable() {
        const string html = "<table><caption>Outer</caption><tr><th>Name</th></tr><tr><td>Parent"
            + "<table><caption>Inner</caption><tr><th>Code</th></tr><tr><td>Nested</td></tr></table>"
            + "</td></tr></table>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "nested.html");

        Assert.Equal(2, result.Tables.Count);
        ReaderTable outer = Assert.Single(result.Tables, table => table.Title == "Outer");
        ReaderTable inner = Assert.Single(result.Tables, table => table.Title == "Inner");
        Assert.Equal(1, outer.TotalRowCount);
        Assert.Equal(1, inner.TotalRowCount);
        Assert.Equal("Nested", inner.Rows[0][0]);
    }

    [Fact]
    public void DocumentReaderHtml_RichTables_ApplyRowLimitToTableBlocks() {
        const string html = "<table><tr><th>Name</th></tr><tr><td>Row 1</td></tr>"
            + "<tr><td>Row 2</td></tr><tr><td>Row 3</td></tr></table>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(
            html,
            "bounded.html",
            new ReaderOptions { MaxTableRows = 1 });

        ReaderTable table = Assert.Single(result.Tables);
        Assert.Single(table.Rows);
        Assert.True(table.Truncated);
        OfficeDocumentBlock block = Assert.Single(result.Blocks, item => item.Kind == "table");
        Assert.Contains("Row 1", block.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 2", block.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 3", block.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_MapsTextareaAndResponsiveImageSources() {
        const string html = "<form><textarea name=\"comments\">Hello there</textarea></form>"
            + "<img alt=\"Responsive\" srcset=\"small.png 1x, large.png 2x\"/>"
            + "<img alt=\"Lazy\" data-src=\"lazy.png\"/>"
            + "<picture><source srcset=\"picture.png 1x\"/></picture>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "responsive.html");

        OfficeDocumentFormField form = Assert.Single(result.Forms, item => item.Name == "comments");
        Assert.Equal("Hello there", form.Value);
        Assert.Contains(result.Assets, asset => asset.SourceObjectId == "small.png");
        Assert.Contains(result.Assets, asset => asset.SourceObjectId == "lazy.png");
        Assert.Contains(result.Assets, asset => asset.SourceObjectId == "picture.png");
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_MapsSelectAsOneFormField() {
        const string html = "<form><select name=\"status\"><option value=\"draft\">Draft</option>"
            + "<option value=\"approved\" selected>Approved</option></select></form>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "select.html");

        OfficeDocumentFormField form = Assert.Single(result.Forms);
        Assert.Equal("status", form.Name);
        Assert.Equal("select", form.Kind);
        Assert.Equal("approved", form.Value);
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_PreservesCheckedStateForCheckboxesAndRadios() {
        const string html = "<form>"
            + "<input type=\"checkbox\" name=\"unchecked\" value=\"yes\">"
            + "<input type=\"checkbox\" name=\"checked\" value=\"yes\" checked>"
            + "<input type=\"radio\" name=\"choice\" value=\"a\">"
            + "<input type=\"radio\" name=\"choice\" value=\"b\" checked>"
            + "</form>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "controls.html");

        Assert.Null(Assert.Single(result.Forms, form => form.Name == "unchecked").Value);
        Assert.Equal("yes", Assert.Single(result.Forms, form => form.Name == "checked").Value);
        Assert.Null(result.Forms.Single(form => form.Kind == "radio" && form.Value == null).Value);
        Assert.Equal("b", Assert.Single(result.Forms, form => form.Kind == "radio" && form.Value != null).Value);
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_PreservesHeaderlessTableRows() {
        const string html = "<table><tr><td>A</td></tr><tr><td>B</td></tr></table>";

        ReaderTable table = Assert.Single(HtmlReaderAdapter.ReadContentDocument(html, "headerless.html").Tables);

        Assert.Equal(new[] { "Column 1" }, table.Columns);
        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal(new[] { "A", "B" }, table.Rows.Select(row => row[0]));
    }

    [Fact]
    public void DocumentReaderHtml_RichProjection_AppliesConfiguredElementFilters() {
        const string html = "<section class=\"private\"><p>Private text</p><a href=\"https://private.test\">Private link</a></section>"
            + "<p id=\"delegate-private\">Delegate private</p><p>Public text</p>";
        var markdownOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        markdownOptions.ExcludeSelectors.Add(".private");
        int delegateMatchCount = 0;
        markdownOptions.ElementFilters.Add(element => {
            if (!string.Equals(element.Id, "delegate-private", StringComparison.Ordinal)) return false;
            delegateMatchCount++;
            return true;
        });
        var options = new ReaderHtmlOptions { HtmlToMarkdownOptions = markdownOptions };

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(
            html,
            "filtered.html",
            htmlOptions: options);

        Assert.DoesNotContain(result.Blocks, block => block.Text.Contains("Private", StringComparison.OrdinalIgnoreCase));
        Assert.Empty(result.Links);
        Assert.DoesNotContain("Private text", result.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("Delegate private", result.Html, StringComparison.Ordinal);
        Assert.Contains(result.Blocks, block => block.Text.Contains("Public text", StringComparison.Ordinal));
        Assert.Equal(1, delegateMatchCount);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_EmitsChunks() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p></body></html>";

        var chunks = HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "inline.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            c.Kind == ReaderInputKind.Html &&
            string.Equals(c.Location.Path, "inline.html", StringComparison.OrdinalIgnoreCase) &&
            ((c.Markdown ?? c.Text).Contains("Hello HTML", StringComparison.Ordinal) ||
             (c.Markdown ?? c.Text).Contains("Body text.", StringComparison.Ordinal)));
        Assert.All(chunks, c => {
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(Encoding.UTF8.GetByteCount(html), c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_TrimsLogicalSourceName() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p></body></html>";

        var chunk = Assert.Single(HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: " inline.html ",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }));

        Assert.Equal("inline.html", chunk.Location.Path);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_UsesHeadingAwareLocations() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p><h2>Second</h2><p>More.</p></body></html>";

        var chunks = HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "headings.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c => !string.IsNullOrWhiteSpace(c.Location.HeadingPath));
        Assert.All(chunks, c => Assert.True(c.Location.StartLine.GetValueOrDefault() >= 1));
        Assert.Contains(chunks, c => (c.Location.HeadingPath ?? string.Empty).Contains("Hello HTML", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderHtml_PreservesLiteralHeadingDisplayAndHierarchy() {
        const string title = "Q1 > Q2\\Back";
        var chunks = HtmlReaderAdapter.ReadContent(
            html: "<html><body><h1>Q1 &gt; Q2\\Back</h1><p>Body.</p></body></html>",
            sourceName: "literal-heading.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        ReaderChunk chunk = Assert.Single(chunks);
        Assert.Equal(title, chunk.Location.HeadingPath);

        ReaderChunkHierarchyResult hierarchy = ReaderHierarchicalChunker.Chunk(chunks);
        ReaderChunkHierarchyNode heading = Assert.Single(
            hierarchy.Nodes,
            node => node.Kind == ReaderChunkHierarchyNodeKind.Heading);
        Assert.Equal(title, heading.Title);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_CanDisableHeadingChunking() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p><h2>Second</h2><p>More.</p></body></html>";

        var chunks = HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "headings-disabled.html",
            readerOptions: new ReaderOptions {
                MaxChars = 8_000
            },
            htmlOptions: new ReaderHtmlOptions { ChunkByHeadings = false }).ToList();

        Assert.NotEmpty(chunks);
        Assert.DoesNotContain(chunks, c => !string.IsNullOrWhiteSpace(c.Location.HeadingPath));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_SplitsByMaxChars() {
        var largeHtml = "<html><body><p>" + new string('a', 2048) + "</p></body></html>";

        var chunks = HtmlReaderAdapter.ReadContent(
            html: largeHtml,
            sourceName: "large.html",
            readerOptions: new ReaderOptions { MaxChars = 128 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Html, c.Kind));
        Assert.Contains(chunks, c =>
            c.Warnings != null &&
            c.Warnings.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_HeadingSplits_DoNotEmitMaxCharsWarning() {
        var html = "<html><body><h1>One</h1><p>Alpha.</p><h1>Two</h1><p>Beta.</p></body></html>";

        var chunks = HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "headings-only.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.DoesNotContain(chunks, c =>
            c.Warnings?.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)) ?? false);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_ExtractsTableProfilesFromConvertedMarkdown() {
        var html =
            "<html><body><h1>Inventory</h1><table>" +
            "<thead><tr><th>Name</th><th>Qty</th></tr></thead>" +
            "<tbody><tr><td>Paper</td><td>10</td></tr><tr><td>Ink</td><td>2</td></tr></tbody>" +
            "</table></body></html>";

        var chunk = HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "inventory.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).Single(c => (c.Tables?.Count ?? 0) > 0);

        Assert.Equal(ReaderInputKind.Html, chunk.Kind);
        Assert.NotNull(chunk.Tables);
        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Name", "Qty" }, table.Columns);
        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal("Paper", table.Rows[0][0]);
        Assert.Equal("2", table.Rows[1][1]);
        Assert.Equal(2, table.ColumnProfiles.Count);
        Assert.Equal(ReaderTableColumnKind.Text, table.ColumnProfiles[0].Kind);
        Assert.Equal(ReaderTableColumnKind.Numeric, table.ColumnProfiles[1].Kind);
        Assert.True(table.ColumnProfiles[1].IsNumeric);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_ExtractsLargeTableProfilesBeforeSplitting() {
        string rows = string.Concat(Enumerable.Range(1, 40).Select(index =>
            "<tr><td>Item " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "</td><td>" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "</td></tr>"));
        string html =
            "<html><body><h1>Inventory</h1><table>" +
            "<thead><tr><th>Name</th><th>Qty</th></tr></thead>" +
            "<tbody>" + rows + "</tbody>" +
            "</table></body></html>";

        var chunks = HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "large-table.html",
            readerOptions: new ReaderOptions { MaxChars = 128 }).ToList();

        Assert.True(chunks.Count > 1);
        var table = Assert.Single(chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>()));
        Assert.Equal(new[] { "Name", "Qty" }, table.Columns);
        Assert.Equal(40, table.TotalRowCount);
        Assert.Equal("Item 1", table.Rows[0][0]);
        Assert.Equal("40", table.Rows[39][1]);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_EmitsWarningForEmptyContent() {
        var chunks = HtmlReaderAdapter.ReadContent(
            html: "<html><body></body></html>",
            sourceName: "empty.html").ToList();

        Assert.Single(chunks);
        var warning = chunks[0];
        Assert.Equal("html-warning-0000", warning.Id);
        Assert.Equal(ReaderInputKind.Html, warning.Kind);
        Assert.Contains("no markdown text", warning.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.False(string.IsNullOrWhiteSpace(warning.SourceId));
        Assert.False(string.IsNullOrWhiteSpace(warning.SourceHash));
        Assert.False(string.IsNullOrWhiteSpace(warning.ChunkHash));
        Assert.True(warning.TokenEstimate.HasValue && warning.TokenEstimate.Value >= 1);
        Assert.Equal(Encoding.UTF8.GetByteCount("<html><body></body></html>"), warning.SourceLengthBytes);
        Assert.Null(warning.SourceLastWriteUtc);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_PreservesConfiguredMaxInputCharacters() {
        var html = "<html><body><p>Too much content for this limit.</p></body></html>";

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "limited.html",
            htmlOptions: new ReaderHtmlOptions {
                HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                    MaxInputCharacters = 12
                }
            }).ToList());

        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderHtmlOptions_ProfilesExposeExpectedMarkdownOptions() {
        var officeProfile = ReaderHtmlOptions.CreateOfficeIMOProfile();
        Assert.NotNull(officeProfile.HtmlToMarkdownOptions);
        Assert.Null(officeProfile.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.Null(officeProfile.HtmlToMarkdownOptions.MaxInputCharacters);

        var portableProfile = ReaderHtmlOptions.CreatePortableProfile();
        Assert.NotNull(portableProfile.HtmlToMarkdownOptions);
        Assert.NotNull(portableProfile.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.Null(portableProfile.HtmlToMarkdownOptions.MaxInputCharacters);

        var untrustedProfile = ReaderHtmlOptions.CreateUntrustedHtmlProfile(64);
        Assert.NotNull(untrustedProfile.HtmlToMarkdownOptions);
        Assert.NotNull(untrustedProfile.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.Equal(64, untrustedProfile.HtmlToMarkdownOptions.MaxInputCharacters);
        Assert.NotNull(untrustedProfile.ConversionOptions);
        Assert.Equal(64, untrustedProfile.ConversionOptions.Limits.MaxInputCharacters);

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => ReaderHtmlOptions.CreateUntrustedHtmlProfile(0));
        Assert.Equal("maxInputCharacters", exception.ParamName);
    }

    [Fact]
    public void ReaderHtmlOptions_CloneCopiesNestedOptionsIndependently() {
        var options = ReaderHtmlOptions.CreateUntrustedHtmlProfile(64);
        options.HtmlToMarkdownOptions!.BaseUri = new Uri("https://example.com/docs/");

        var clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.NotNull(clone.HtmlToMarkdownOptions);
        Assert.NotSame(options.HtmlToMarkdownOptions, clone.HtmlToMarkdownOptions);
        Assert.Equal(options.HtmlToMarkdownOptions.BaseUri, clone.HtmlToMarkdownOptions.BaseUri);
        Assert.Equal(options.HtmlToMarkdownOptions.MaxInputCharacters, clone.HtmlToMarkdownOptions.MaxInputCharacters);
        Assert.NotNull(clone.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.NotSame(options.HtmlToMarkdownOptions.MarkdownWriteOptions, clone.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.NotNull(clone.ConversionOptions);
        Assert.NotSame(options.ConversionOptions, clone.ConversionOptions);
        Assert.NotSame(options.ConversionOptions!.Limits, clone.ConversionOptions.Limits);

        clone.HtmlToMarkdownOptions.MaxInputCharacters = 128;
        clone.ConversionOptions.Limits.MaxInputCharacters = 128;

        Assert.Equal(64, options.HtmlToMarkdownOptions.MaxInputCharacters);
        Assert.Equal(128, clone.HtmlToMarkdownOptions.MaxInputCharacters);
        Assert.Equal(64, options.ConversionOptions.Limits.MaxInputCharacters);
        Assert.Equal(128, clone.ConversionOptions.Limits.MaxInputCharacters);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_UntrustedProfileEnforcesMaxInputCharacters() {
        var html = "<html><body><p>Too much content for this profile limit.</p></body></html>";

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => HtmlReaderAdapter.ReadContent(
            html: html,
            sourceName: "profile-limited.html",
            htmlOptions: ReaderHtmlOptions.CreateUntrustedHtmlProfile(12)).ToList());

        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlStream_EmitsLogicalSourceMetadata() {
        var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);

        var chunks = HtmlReaderAdapter.Read(
            stream,
            sourceName: " metadata.html ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => {
            Assert.Equal("metadata.html", c.Location.Path);
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(stream.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlFile_EmitsFileSourceMetadata() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-html-meta-" + Guid.NewGuid().ToString("N") + ".html");
        try {
            File.WriteAllText(path, "<html><body><h1>File HTML</h1><p>Body text.</p></body></html>");

            var chunks = HtmlReaderAdapter.Read(
                path,
                readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 8_000 }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => {
                Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
                Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
                Assert.True(c.SourceLengthBytes.HasValue && c.SourceLengthBytes.Value > 0);
                Assert.True(c.SourceLastWriteUtc.HasValue);
            });
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderHtml_BuilderHandler_DispatchesHtmlStream() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();

        var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
        var chunks = reader.Read(stream, "registry.html").ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            c.Kind == ReaderInputKind.Html &&
            string.Equals(c.Location.Path, "registry.html", StringComparison.OrdinalIgnoreCase) &&
            ((c.Markdown ?? c.Text).Contains("Registry HTML", StringComparison.Ordinal) ||
             (c.Markdown ?? c.Text).Contains("From stream.", StringComparison.Ordinal)));
    }

    [Fact]
    public void DocumentReaderHtml_BuilderHandler_DispatchesXhtmlStream() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();

        var html = "<html><body><h2>Registry XHTML</h2><p>From stream.</p></body></html>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
        var chunks = reader.Read(stream, "registry.xhtml").ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            c.Kind == ReaderInputKind.Html &&
            string.Equals(c.Location.Path, "registry.xhtml", StringComparison.OrdinalIgnoreCase) &&
            ((c.Markdown ?? c.Text).Contains("Registry XHTML", StringComparison.Ordinal) ||
             (c.Markdown ?? c.Text).Contains("From stream.", StringComparison.Ordinal)));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlStream_NonSeekable_EnforcesMaxInputBytes() {
        var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(html));

        var ex = Assert.Throws<IOException>(() => HtmlReaderAdapter.Read(
            stream,
            sourceName: "nonseekable.html",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlFile_EnforcesMaxInputBytes() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-html-" + Guid.NewGuid().ToString("N") + ".html");
        try {
            File.WriteAllText(path, "<html><body><p>" + new string('a', 256) + "</p></body></html>");

            var ex = Assert.Throws<IOException>(() => HtmlReaderAdapter.Read(
                path,
                readerOptions: new ReaderOptions { MaxInputBytes = 32 }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderHtml_BuilderHandler_NonSeekableStream_EnforcesMaxInputBytes() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();

        var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(html));
        var ex = Assert.Throws<IOException>(() => reader.Read(
            stream,
            "registry.html",
            new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_BuilderHandler_PreservesConfiguredMaxInputCharacters() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHtmlHandler(new ReaderHtmlOptions {
                    HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                        MaxInputCharacters = 12
                    }
                })
            .Build();

        var html = "<html><body><p>Too much content for this limit.</p></body></html>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => reader.Read(stream, "configured.html").ToList());

        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_BuilderHandler_AppliesConfiguredMarkdownOptions() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHtmlHandler(new ReaderHtmlOptions {
                    HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                        BaseUri = new Uri("https://example.com/docs/")
                    }
                })
            .Build();

        var html = "<html><body><p><a href=\"guide/getting-started\">Docs</a></p></body></html>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
        var chunks = reader.Read(stream, "configured.html").ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            ((c.Markdown ?? c.Text).Contains("https://example.com/docs/guide/getting-started", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void DocumentReaderHtml_ProjectsAccessibilityAndFootnoteCapabilitiesThroughJson() {
        const string html = """
<html>
  <head><base href="https://example.test/book/"></head>
  <body>
    <div role="heading" aria-level="4">Accessible section</div>
    <p><a href="note.xhtml" aria-label="Open note">Visible note</a></p>
    <blockquote><p>Quoted text</p></blockquote>
    <pre data-language="csharp">Console.WriteLine(1);</pre>
    <ol type="A" start="3"><li>Third choice</li></ol>
    <span id="cover-label">Cover chart</span>
    <img src="images/cover.png" aria-labelledby="cover-label">
    <aside id="note" epub:type="footnote" role="doc-footnote"><p>Footnote body</p></aside>
  </body>
</html>
""";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "chapter.xhtml");

        OfficeDocumentBlock heading = Assert.Single(result.Blocks, block => block.Kind == "heading");
        OfficeDocumentLink link = Assert.Single(result.Links);
        OfficeDocumentAsset image = Assert.Single(result.Assets);
        ReaderVisual visual = Assert.Single(result.Visuals, item => item.Kind == "image");
        OfficeDocumentBlock quote = Assert.Single(result.Blocks, block => block.Kind == "quote");
        OfficeDocumentBlock code = Assert.Single(result.Blocks, block => block.Kind == "code");
        OfficeDocumentBlock footnote = Assert.Single(result.Blocks, block => block.Kind == "footnote");
        OfficeDocumentBlock listItem = Assert.Single(result.Blocks, block => block.Kind == "list-item");

        Assert.Equal(4, heading.Level);
        Assert.Equal("Accessible section", heading.Text);
        Assert.Equal("Open note", link.Text);
        Assert.Equal("https://example.test/book/note.xhtml", link.Uri);
        Assert.Equal("Cover chart", image.AltText);
        Assert.Equal("https://example.test/book/images/cover.png", image.SourceObjectId);
        Assert.Equal("Cover chart", visual.Content);
        Assert.Equal("Quoted text", quote.Text);
        Assert.Equal("Console.WriteLine(1);", code.Text);
        Assert.Equal("Footnote body", footnote.Text);
        Assert.Equal("C.", listItem.Marker);
        Assert.Contains("officeimo.html.accessibility", result.CapabilitiesUsed);
        Assert.Contains("officeimo.html.footnotes", result.CapabilitiesUsed);
        Assert.Contains("officeimo.html.quotes", result.CapabilitiesUsed);
        Assert.Contains("officeimo.html.code", result.CapabilitiesUsed);

        OfficeDocumentReadResult roundTrip = OfficeDocumentReadResultJson.Deserialize(
            HtmlReaderAdapter.ReadContentDocumentJson(html, "chapter.xhtml"));
        Assert.Equal("Cover chart", Assert.Single(roundTrip.Assets).AltText);
        Assert.Contains("officeimo.html.accessibility", roundTrip.CapabilitiesUsed);
        Assert.Contains("officeimo.html.footnotes", roundTrip.CapabilitiesUsed);
    }

    [Fact]
    public void DocumentReaderHtml_UsesAriaColumnHeadersForTableColumns() {
        const string html = """
<div role="table">
  <div role="row"><div role="columnheader">Name</div><div role="columnheader">Value</div></div>
  <div role="row"><div role="cell">EPUB</div><div role="cell">Structured</div></div>
</div>
""";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "aria-table.html");
        ReaderTable table = Assert.Single(result.Tables);

        Assert.Equal(new[] { "Name", "Value" }, table.Columns);
        Assert.Equal(new[] { "EPUB", "Structured" }, Assert.Single(table.Rows));
    }

    [Fact]
    public void DocumentReaderHtml_DoesNotDuplicateNestedQuoteOrFootnoteBlocks() {
        const string html = """
<blockquote>
  <h2>Quoted heading</h2>
  <pre>quoted code</pre>
  <ul><li>quoted item</li></ul>
  <table><tr><th>Quoted column</th></tr><tr><td>quoted value</td></tr></table>
</blockquote>
<aside epub:type="footnote" id="source-note">
  <p>Footnote text</p>
  <pre>footnote code</pre>
</aside>
""";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "nested-blocks.html");
        OfficeDocumentBlock quote = Assert.Single(result.Blocks, block => block.Kind == "quote");
        OfficeDocumentBlock footnote = Assert.Single(result.Blocks, block => block.Kind == "footnote");

        Assert.Contains("Quoted heading", quote.Text, StringComparison.Ordinal);
        Assert.Contains("quoted code", quote.Text, StringComparison.Ordinal);
        Assert.Contains("Footnote text", footnote.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Blocks, block => block.Kind == "code" || block.Kind == "list-item" || block.Kind == "table");
        Assert.Empty(result.Tables);
    }

    [Fact]
    public void DocumentReaderHtml_DoesNotPromoteGenericEpubNotesToFootnotes() {
        const string html = "<aside epub:type=\"note\" id=\"editorial\"><p>Editorial context.</p></aside>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "generic-note.html");

        Assert.DoesNotContain(result.Blocks, block => block.Kind == "footnote");
        Assert.DoesNotContain("officeimo.html.footnotes", result.CapabilitiesUsed);
        Assert.Contains(result.Blocks, block => block.Kind == "paragraph" && block.Text == "Editorial context.");
    }
}
