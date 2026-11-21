using System.Management.Automation;
using System.Management.Automation.Runspaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Cmdlets.Word;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class WordFieldUpdateWorkflowTests {
    [Fact]
    public void TableOfContentsTakesPageNumbersFromTheGivenPdfLayout() {
        var pdf = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using WordDocument document = WordDocument.Create();
        document.AddTableOfContent();
        document.AddPageBreak();
        document.AddParagraph("Overview").Style = WordParagraphStyles.Heading1;
        document.AddPageBreak();
        document.AddParagraph("Details").Style = WordParagraphStyles.Heading1;

        using Runspace runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Update-OfficeWordField").AddParameter("Document", document)
            .AddParameter("TableOfContents").AddParameter("LayoutOptions", pdf);
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));

        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var pages = package.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>().Where(link => link.Anchor?.Value != null)
            .ToDictionary(link => link.Descendants<Text>().First().Text, link => link.Descendants<Text>().Last().Text);
        Assert.Equal(new Dictionary<string, string> { ["Overview"] = "2", ["Details"] = "3" }, pages);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Update-OfficeWordField", typeof(UpdateOfficeWordFieldCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
