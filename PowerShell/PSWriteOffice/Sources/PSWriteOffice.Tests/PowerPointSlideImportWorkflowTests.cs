using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointSlideImportWorkflowTests {
    [Theory]
    [InlineData("Import-OfficePowerPointSlide -Document $deck -SourceDocument $library -SourceSlide 3, 1 -Position 2 -PassThru")]
    [InlineData("$deck | Import-OfficePowerPointSlide -SourceDocument $library -SourceSlide 3, 1 -Position 2 -PassThru")]
    [InlineData("Import-OfficePowerPointSlide -Document $deck -LiteralPath $path -SourceSlide 3, 1 -Position 2 -PassThru")]
    [InlineData("Get-Item -LiteralPath $path | Import-OfficePowerPointSlide -Document $deck -SourceSlide 3, 1 -Position 2 -PassThru")]
    public void ImportUsesOneBasedSelectionAndReturnsEditableSavedSlides(string command) {
        WithDocuments((shell, deck, library, _) => {
            shell.AddScript(command);
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Equal(2, result.Count);
            Assert.Same(deck.Slides[1], result[0].BaseObject);
            Assert.Same(deck.Slides[2], result[1].BaseObject);
            deck.Slides[1].AddTextBox("Edited after import");
            using var saved = PowerPointDocument.Load(new MemoryStream(deck.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            Assert.Equal(new[] { "Target", "Third", "First" }, saved.Slides.Select(slide => slide.TextBoxes.First().Text));
            Assert.Contains(saved.Slides[1].TextBoxes, box => box.Text == "Edited after import");
            Assert.Equal("Source notes", saved.Slides[2].Notes.Text);
            Assert.Equal(3, library.Slides.Count);
        });
    }

    [Theory]
    [InlineData("Import-OfficePowerPointSlide -Document $deck -SourceDocument $library", false)]
    [InlineData("Import-OfficePowerPointSlide -Document $deck -Path $wildcard", false)]
    [InlineData("Import-OfficePowerPointSlide -Document $deck -Path $wildcard -Position 2", true)]
    public void OmittedSelectionImportsTheWholeSourceInFileOrderWithoutPipelineOutput(string command, bool multipleFiles) {
        WithDocuments((shell, deck, _, path) => {
            if (multipleFiles) {
                using var appendix = PowerPointDocument.Create();
                appendix.AddSlide().AddTextBox("Appendix");
                appendix.Save(Path.Combine(Path.GetDirectoryName(path)!, "Library~appendix.pptx"));
            }
            shell.Runspace.SessionStateProxy.SetVariable("wildcard", Path.Combine(Path.GetDirectoryName(path)!, "Library*.pptx"));
            shell.AddScript(command);
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
            using var saved = PowerPointDocument.Load(new MemoryStream(deck.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            string[] expected = multipleFiles ? new[] { "Target", "First", "Second", "Third", "Appendix" } : new[] { "Target", "First", "Second", "Third" };
            Assert.Equal(expected, saved.Slides.Select(slide => slide.TextBoxes.First().Text));
        });
    }

    [Theory]
    [InlineData("-SourceSlide 1, 99", "sourceIndices")]
    [InlineData("-SourceSlide 1 -Position 3", "insertAt")]
    public void InvalidSelectionOrPositionDoesNotPartiallyImport(string options, string parameter) {
        WithDocuments((shell, deck, _, _) => {
            var original = deck.Slides[0];
            shell.AddScript("Import-OfficePowerPointSlide -Document $deck -SourceDocument $library " + options);
            Exception? error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            var failure = error?.GetBaseException() ?? Assert.Single(shell.Streams.Error).Exception.GetBaseException();
            Assert.Equal(parameter, Assert.IsType<ArgumentOutOfRangeException>(failure).ParamName);
            Assert.Same(original, Assert.Single(deck.Slides));
        });
    }

    private static void WithDocuments(Action<PowerShell, PowerPointDocument, PowerPointDocument, string> action) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-import-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Library[final].pptx");
            using var library = PowerPointDocument.Create();
            foreach (string title in new[] { "First", "Second", "Third" }) library.AddSlide().AddTextBox(title);
            library.Slides[0].Notes.Text = "Source notes";
            library.Save(path);
            using var deck = PowerPointDocument.Create();
            deck.AddSlide().AddTextBox("Target");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Import-OfficePowerPointSlide", typeof(ImportOfficePowerPointSlideCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("deck", deck);
            runspace.SessionStateProxy.SetVariable("library", library);
            runspace.SessionStateProxy.SetVariable("path", path);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            action(shell, deck, library, path);
        } finally { Directory.Delete(root, true); }
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
