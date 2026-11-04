using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Word;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class WordHeadingWorkflowTests {
    [Fact]
    public void TextAndRichHeadingsRetainTheirNativeStylesAfterSaving() {
        using var document = WordDocument.Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("document", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $document | WordHeading 'Review' -Level Title
            WordHeading -Document $document 'Overview' -Level 2 -Alignment Center -PassThru
            WordHeading -Document $document -Level 3 { WordText 'Incidents '; WordText '(2)' -Bold }
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.IsType<WordParagraph>(Assert.Single(result).BaseObject);
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.Equal(WordParagraphStyles.Heading2, reopened.Paragraphs[1].Style);
        Assert.Equal(WordParagraphStyles.Heading3, reopened.Paragraphs[2].Style);
        Assert.Equal(WordParagraphAlignment.Center, reopened.Paragraphs[1].ParagraphAlignment);
        Assert.True(reopened.Paragraphs[2].GetRuns().Last().Bold);
        using var packageBytes = new MemoryStream(saved.ToArray());
        using var package = WordprocessingDocument.Open(packageBytes, false);
        // Word's paragraph collection exposes run wrappers; the package owns logical paragraph boundaries.
        var paragraphs = package.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToArray();
        Assert.Equal(new[] { "Review", "Overview", "Incidents (2)" }, paragraphs.Select(p => p.InnerText));
        Assert.Equal(new[] { "Title", "Heading2", "Heading3" },
            paragraphs.Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value));
        Assert.NotNull(paragraphs[2].Elements<Run>().Last().RunProperties?.Bold);
    }

    [Fact]
    public void InvalidHeadingLevelDoesNotAppendAParagraph() {
        using var document = WordDocument.Create();
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Add-OfficeWordHeading").AddParameter("Document", document)
            .AddParameter("Text", "Invalid").AddParameter("Level", 99);
        var failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Empty(document.Paragraphs);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        foreach (Type type in typeof(AddOfficeWordHeadingCommand).Assembly.GetTypes()) {
            var command = type.GetCustomAttribute<CmdletAttribute>();
            if (command == null || !command.NounName.StartsWith("OfficeWord", StringComparison.Ordinal)) continue;
            string name = command.VerbName + "-" + command.NounName;
            state.Commands.Add(new SessionStateCmdletEntry(name, type, null));
            foreach (string alias in type.GetCustomAttribute<AliasAttribute>()?.AliasNames ?? Array.Empty<string>())
                state.Commands.Add(new SessionStateAliasEntry(alias, name));
        }
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
