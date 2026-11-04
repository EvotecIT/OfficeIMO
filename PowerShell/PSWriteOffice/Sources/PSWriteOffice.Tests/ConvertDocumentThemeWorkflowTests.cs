using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Workflows;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ConvertDocumentThemeWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-convert-theme-" + Guid.NewGuid().ToString("N"));
    public ConvertDocumentThemeWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, recursive: true);

    [Theory]
    [InlineData("Path", "Markdown", ".docx", "Name", null)]
    [InlineData("PathFile", "Html", ".docx", "Object", null)]
    [InlineData("LiteralPath", "Markdown", ".pptx", "Potx", null)]
    [InlineData("LiteralPathFile", "Html", ".docx", "Dotx", null)]
    [InlineData("Document", "Markdown", ".pptx", "Object", null)]
    [InlineData("Text", "Text", ".pptx", "Name", null)]
    [InlineData("Text", "Text", ".docx", "Name", "markdown-docx")]
    public void CreationRoutesApplyTheOverrideWithoutMutatingCallerOptions(string parameterSet, string from, string extension, string themeKind, string? route) {
        OfficeTheme expected = OfficeTheme.Boardroom;
        if (themeKind == "Object") expected = OfficeTheme.Classic.WithName("Custom conversion theme");
        object themeInput = themeKind == "Name" ? "bOaRdRoOm" : expected;
        string? template = null;
        if (themeKind is "Dotx" or "Potx") {
            template = Path.Combine(_root, "Corporate." + themeKind.ToLowerInvariant());
            if (themeKind == "Dotx") {
                using (var source = WordDocument.Create(new WordCreateOptions { Theme = expected })) {
                    source.AddParagraph("Template content must stay behind");
                    source.Save(template);
                }
                using var package = WordprocessingDocument.Open(template, true);
                package.ChangeDocumentType(WordprocessingDocumentType.Template);
            } else {
                using (var source = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = expected })) {
                    source.AddSlide().AddTextBox("Template content must stay behind");
                    source.Save(template);
                }
                using var package = PresentationDocument.Open(template, true);
                package.ChangeDocumentType(PresentationDocumentType.Template);
            }
            // Reading a template retains its palette and fonts, with Modern composition defaults.
            expected = OfficeTheme.FromDocument(template);
            themeInput = template;
        }
        byte[]? templateBytes = template == null ? null : File.ReadAllBytes(template);
        string content = from == "Html" ? "<h1>Conversion heading</h1><p>Body evidence.</p>" : "# Conversion heading\n\nBody evidence.";
        string sourcePath = Path.Combine(_root, "source" + (from == "Html" ? ".html" : ".md"));
        File.WriteAllText(sourcePath, content);
        string output = Path.Combine(_root, "output", "source" + extension);
        var wordOptions = new MarkdownWordExportOptions { Theme = MarkdownVisualTheme.FromOfficeTheme(OfficeTheme.Modern) };
        var htmlOptions = new HtmlWordExportOptions { Theme = OfficeTheme.Modern };
        var deckOptions = new MarkdownToPowerPointOptions { Theme = OfficeTheme.Modern };
        object options = extension == ".pptx" ? deckOptions : from == "Html" ? htmlOptions : wordOptions;
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument");
        using var markdown = MarkdownDocument.Parse(content);
        switch (parameterSet) {
            case "Text":
                shell.AddParameter("Text", content).AddParameter("From", from);
                break;
            case "Document":
                shell.AddParameter("Document", markdown);
                break;
            default:
                shell.AddParameter(parameterSet.StartsWith("Literal", StringComparison.Ordinal) ? "LiteralPath" : "Path", sourcePath);
                break;
        }
        if (parameterSet is "Path" or "LiteralPath")
            shell.AddParameter("OutputDirectory", Path.GetDirectoryName(output));
        else shell.AddParameter("OutputPath", output);
        shell.AddParameter("To", extension).AddParameter("Theme", themeInput).AddParameter("Options", new[] { options })
            .AddParameter("ErrorAction", ActionPreference.Stop);
        if (route != null) shell.AddParameter("Route", route);
        var result = Assert.IsType<OfficeConversionReport>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(output, result.OutputPath);
        Assert.Same(OfficeTheme.Modern, wordOptions.Theme!.SourceTheme);
        Assert.Same(OfficeTheme.Modern, htmlOptions.Theme);
        Assert.Same(OfficeTheme.Modern, deckOptions.Theme);
        Assert.Equal(content, File.ReadAllText(sourcePath));
        if (template != null) Assert.Equal(templateBytes, File.ReadAllBytes(template));
        AssertThemeAndContent(output, expected);
    }

    [Fact]
    public void ThemeRequiresACreationRouteInTheActualFolderSelection() {
        string input = Path.Combine(_root, "input");
        Directory.CreateDirectory(input);
        File.WriteAllText(Path.Combine(input, "page.html"), "<p>Not a deck source</p>");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", input).AddParameter("To", "PowerPoint")
            .AddParameter("OutputDirectory", Path.Combine(_root, "output")).AddParameter("Theme", "Boardroom")
            .AddParameter("ErrorAction", ActionPreference.Stop);
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        string error = failure?.Message ?? string.Join(Environment.NewLine, shell.Streams.Error);
        Assert.Contains("Theme requires a selected", error);
        Assert.False(Directory.Exists(Path.Combine(_root, "output")));
    }

    [Fact]
    public void SavedJobsRejectExecutionOnlyThemesBeforeWritingADefinition() {
        string input = Path.Combine(_root, "source.md");
        File.WriteAllText(input, "# Saved source");
        string job = Path.Combine(_root, "conversion.officejob.json");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", input).AddParameter("To", "Word")
            .AddParameter("OutputDirectory", Path.Combine(_root, "output")).AddParameter("Theme", OfficeTheme.Classic)
            .AddParameter("SaveJob", job).AddParameter("ErrorAction", ActionPreference.Stop);
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Contains("portable scalar settings", failure?.Message ?? string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.False(File.Exists(job));
        Assert.False(Directory.Exists(Path.Combine(_root, "output")));
        // The same creation route remains portable when no execution-only options are supplied.
        shell.Commands.Clear();
        shell.Streams.Error.Clear();
        shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", input).AddParameter("To", "Word")
            .AddParameter("OutputDirectory", Path.Combine(_root, "output")).AddParameter("SaveJob", job)
            .AddParameter("ErrorAction", ActionPreference.Stop);
        Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Null(OfficeBatchOptions.Load(job).ConversionOptions.MarkdownWord);
    }

    private static void AssertThemeAndContent(string path, OfficeTheme expected) {
        using var package = ZipFile.OpenRead(path);
        var themePart = Assert.Single(package.Entries, entry => entry.FullName.EndsWith("/theme1.xml", StringComparison.Ordinal));
        using var themeStream = themePart.Open();
        XDocument theme = XDocument.Load(themeStream);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        Assert.Equal(expected.Fonts.Heading.Family, (string?)theme.Descendants(drawing + "majorFont").Elements(drawing + "latin").Single().Attribute("typeface"));
        Assert.Equal(expected.Fonts.Body.Family, (string?)theme.Descendants(drawing + "minorFont").Elements(drawing + "latin").Single().Attribute("typeface"));
        Assert.Equal(expected.Palette.Accent1.ToRgbHex(), (string?)theme.Descendants(drawing + "accent1").Elements(drawing + "srgbClr").Single().Attribute("val"));
        string part = path.EndsWith(".pptx", StringComparison.Ordinal) ? "ppt/slides/slide1.xml" : "word/document.xml";
        using var contentStream = package.GetEntry(part)!.Open();
        XDocument content = XDocument.Load(contentStream);
        Assert.Contains("Conversion heading", content.ToString());
        Assert.Contains("Body evidence.", content.ToString());
        Assert.DoesNotContain("Template content must stay behind", content.ToString());
        if (part.StartsWith("ppt/", StringComparison.Ordinal)) {
            XElement body = Assert.Single(content.Descendants(drawing + "r"), run => (string?)run.Element(drawing + "t") == "Body evidence.");
            Assert.Equal(expected.Fonts.Body.Family, (string?)body.Element(drawing + "rPr")?.Element(drawing + "latin")?.Attribute("typeface"));
        }
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
