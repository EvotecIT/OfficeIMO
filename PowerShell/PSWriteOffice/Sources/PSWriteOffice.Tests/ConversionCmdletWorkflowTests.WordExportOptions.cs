using System.Management.Automation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.GoogleDocs;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void WordExportFactoriesPreserveDefaultsAndRouteExplicitSwitchesToActualConversions(bool enabled) {
        string root = Directory.CreateTempSubdirectory("officeimo-word-export-options-").FullName;
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Underlined evidence");
        paragraph.Underline = WordUnderlineStyle.Single;
        document.AddParagraph("Footnote anchor").AddFootNote("Distinct footnote evidence");
        try {
            string images = Path.Combine(root, "images");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            runspace.SessionStateProxy.SetVariable("enabled", enabled);
            runspace.SessionStateProxy.SetVariable("images", images);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                New-OfficeWordHtmlOptions
                $htmlOptions = New-OfficeWordHtmlOptions -IncludeDefaultCss:$enabled -ExcludeFootnotes:$enabled -UseImagePaths:$enabled
                $markdownOptions = New-OfficeWordMarkdownOptions -EnableUnderline:$enabled -EnableHighlight:$false -ImageExportMode File -ImageDirectory $images
                if (Test-Path -LiteralPath $images) { throw 'Option creation wrote an image folder' }
                $document | Convert-OfficeDocument -To Html -Options $htmlOptions -ErrorAction Stop
                $document | Convert-OfficeDocument -To Markdown -Options $markdownOptions -ErrorAction Stop
                $htmlOptions
                $markdownOptions
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(5, output.Count);
            var defaults = Assert.IsType<WordHtmlExportOptions>(output[0].BaseObject);
            Assert.True(defaults.IncludeDefaultCss); Assert.True(defaults.ExportFootnotes); Assert.True(defaults.EmbedImagesAsBase64);
            string html = Assert.IsType<string>(output[1].BaseObject);
            Assert.Contains("Underlined evidence", html, StringComparison.Ordinal);
            Assert.Equal(!enabled, html.Contains("Distinct footnote evidence", StringComparison.Ordinal));
            string markdown = Assert.IsType<string>(output[2].BaseObject);
            Assert.Equal(enabled, markdown.Contains("<u>Underlined evidence</u>", StringComparison.Ordinal));
            var htmlOptions = Assert.IsType<WordHtmlExportOptions>(output[3].BaseObject);
            Assert.Equal(enabled, htmlOptions.IncludeDefaultCss);
            Assert.Equal(!enabled, htmlOptions.ExportFootnotes); Assert.Equal(!enabled, htmlOptions.EmbedImagesAsBase64);
            var markdownOptions = Assert.IsType<WordMarkdownExportOptions>(output[4].BaseObject);
            Assert.Equal(images, markdownOptions.ImageDirectory); Assert.False(markdownOptions.EnableHighlight);
            Assert.Equal(WordMarkdownImageExportMode.File, markdownOptions.ImageExportMode);
            Assert.Equal(WordUnderlineStyle.Single, paragraph.Underline);
            Assert.Contains(paragraph, document.Paragraphs);
        } finally { Directory.Delete(root, true); }
    }
    [Fact]
    public void HtmlWordFactoryResolvesLiteralStylesheetsAndAppliesThemDuringConversion() {
        string root = Directory.CreateTempSubdirectory("officeimo-html-word-options-").FullName;
        try {
            string css = Path.Combine(root, "report[1].css"), outputPath = Path.Combine(root, "report.docx");
            File.WriteAllText(css, ".evidence { color: #112233; }");
            var source = OfficeIMO.Html.HtmlConversionDocument.Parse("<p class='evidence'>Styled evidence</p><pre>Preformatted evidence</pre>");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source); runspace.SessionStateProxy.SetVariable("css", css);
            runspace.SessionStateProxy.SetVariable("outputPath", outputPath);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $options = New-OfficeHtmlWordOptions -StylesheetPath $css -StylesheetContent 'pre { font-family: Consolas; }' -RenderPreAsTable -SectionTagHandling Block
                Convert-OfficeDocument -Document $source -To Word -OutputPath $outputPath -Options $options -ErrorAction Stop | Out-Null
                $options
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            var options = Assert.IsType<WordHtmlImportOptions>(Assert.Single(output).BaseObject);
            Assert.Equal(css, Assert.Single(options.StylesheetPaths));
            Assert.False(options.AllowDocumentStylesheetLinks);
            using var document = WordDocument.Load(outputPath);
            var paragraph = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Styled evidence");
            Assert.Equal("112233", paragraph.ColorHex);
            var table = Assert.Single(document.Tables);
            Assert.Contains(table.Rows[0].Cells[0].Paragraphs, paragraph => paragraph.Text == "Preformatted evidence");
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void GoogleDocsFactoryFeedsOfflineBatchCompilationWithoutCreatingASession() {
        using var document = WordDocument.Create();
        document.AddParagraph("Offline Google batch evidence");
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("document", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $defaults = New-OfficeWordGoogleDocsOptions
            $options = New-OfficeWordGoogleDocsOptions -Title 'Prepared report' -ExistingFileId synthetic-file -TabStrategy SelectedTab -TabId synthetic-tab -ExpectedRevisionId synthetic-revision -Comments Skip
            Export-OfficeWordGoogleDocument -Document $document -Options @($options, @{ Tabs = $options.Tabs }, @{ Title = 'Scoped report'; Tabs = @{ TabId = 'override-tab' } }) -Mode Batch -ErrorAction Stop
            $defaults
            $options
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(3, output.Count);
        var batch = Assert.IsType<GoogleDocsBatch>(output[0].BaseObject);
        Assert.Equal("Scoped report", batch.Title);
        Assert.Contains(batch.Requests.OfType<GoogleDocsInsertParagraphRequest>(), item => item.Paragraph.Text.Contains("Offline Google batch evidence", StringComparison.Ordinal));
        var defaults = Assert.IsType<GoogleDocsSaveOptions>(output[1].BaseObject);
        Assert.Equal(GoogleDocsCommentMode.UnanchoredDriveComments, defaults.Comments);
        Assert.Equal(GoogleDocsRevisionConflictMode.RequireRevision, defaults.Replace.ConflictMode);
        var options = Assert.IsType<GoogleDocsSaveOptions>(output[2].BaseObject);
        Assert.Equal("Prepared report", options.Title);
        Assert.Equal("synthetic-file", options.Location.ExistingFileId);
        Assert.Equal("synthetic-tab", options.Tabs.TabId);
        Assert.Equal(GoogleDocsTabStrategy.SelectedTab, options.Tabs.Strategy);
        Assert.Equal("synthetic-revision", options.Replace.ExpectedRevisionId);
        Assert.Equal(GoogleDocsCommentMode.Skip, options.Comments);
        Assert.Equal(GoogleDocsRevisionConflictMode.RequireRevision, options.Replace.ConflictMode);
        Assert.Equal("Offline Google batch evidence", document.Paragraphs[0].Text);
    }

}
