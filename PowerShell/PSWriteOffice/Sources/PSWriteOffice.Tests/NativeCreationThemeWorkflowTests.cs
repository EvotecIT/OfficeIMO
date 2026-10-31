using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Excel;
using PSWriteOffice.Cmdlets.PowerPoint;
using PSWriteOffice.Cmdlets.Word;
using PSWriteOffice.Cmdlets.Visio;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class NativeCreationThemeWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-new-theme-" + Guid.NewGuid().ToString("N"));
    public NativeCreationThemeWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Theory]
    [InlineData("Word", ".docx")]
    [InlineData("Excel", ".xlsx")]
    [InlineData("PowerPoint", ".pptx")]
    public void ThemePathReadsTheVisualIdentityWithoutCopyingTemplateContent(string format, string extension) {
        string template = Path.Combine(_root, "Corporate.docx");
        using (var source = WordDocument.Create(new WordCreateOptions { Theme = OfficeTheme.Classic })) {
            source.AddParagraph("Content only for the template");
            source.Save(template);
        }
        byte[] original = File.ReadAllBytes(template);
        string path = Path.Combine(_root, "Created" + extension);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        var output = shell.AddCommand("New-Office" + format).AddParameter("Theme", template).AddParameter("Path", path).Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using (var document = Assert.IsAssignableFrom<IDisposable>(Assert.Single(output).BaseObject)) {
            Assert.False(File.Exists(path));
            switch (document) {
                case WordDocument word:
                    Assert.DoesNotContain(word.Paragraphs, paragraph => paragraph.Text.Contains("Content only for the template"));
                    word.Save(); break;
                case ExcelDocument excel:
                    Assert.Empty(excel.Sheets);
                    excel.AddSheet("Data"); excel.Save(); break;
                case PowerPointDocument presentation:
                    Assert.Empty(presentation.Slides);
                    presentation.AddSlide(); presentation.Save(); break;
            }
        }
        using var package = ZipFile.OpenRead(path);
        ZipArchiveEntry themePart = Assert.Single(package.Entries, part => part.FullName.EndsWith("/theme1.xml", StringComparison.Ordinal));
        using Stream input = themePart.Open();
        XDocument theme = XDocument.Load(input);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        Assert.Contains(theme.Descendants(drawing + "minorFont").Elements(drawing + "latin"),
            font => (string?)font.Attribute("typeface") == OfficeTheme.Classic.Fonts.Body.Family);
        Assert.Equal(original, File.ReadAllBytes(template));
    }

    [Theory]
    [InlineData("Word", ".docx")]
    [InlineData("Excel", ".xlsx")]
    [InlineData("Visio", ".vsdx")]
    public void FailedCreationBlockPreservesTheDestinationAndRemovesItsPrivateOutput(string format, string extension) {
        string path = Path.Combine(_root, "Existing" + extension);
        byte[] original = { 1, 2, 3, 4 };
        File.WriteAllBytes(path, original);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        runspace.SessionStateProxy.SetVariable("entered", false);
        shell.AddCommand("New-Office" + format).AddParameter("Path", path).AddParameter("Force")
            .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; throw 'authoring-failed'"));
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True((bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.True(error != null || shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(path, Assert.Single(Directory.GetFiles(_root)));
    }

    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void LiveCreationRejectsSaveOnlyPasswordsBeforeCreatingTheDestination(string format) {
        string path = Path.Combine(_root, "missing", "Output");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-Office" + format).AddParameter("Path", path).AddParameter("Password", "test-only");
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        string message = error?.Message ?? Errors(shell);
        Assert.True(message.Contains("requires a content block") || message.Contains("require a content block"), message);
        Assert.Empty(Directory.GetFileSystemEntries(_root));
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeWord", typeof(NewOfficeWordCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcel", typeof(NewOfficeExcelCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPoint", typeof(NewOfficePowerPointCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeVisio", typeof(NewOfficeVisioCommand), null));
        Runspace runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
