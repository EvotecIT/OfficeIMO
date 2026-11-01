using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Word;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class WordImageWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FileAndPreparedVisualPlacementKeepTypedGeometryMetadataAndNativeEdits(bool preparedVisual) {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"300\" height=\"100\"><rect width=\"300\" height=\"100\" fill=\"#147d92\"/></svg>";
        string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "[1].svg");
        File.WriteAllText(source, svg);
        try {
            var visual = new OfficeVisualSource(svg) { Title = "Original", AlternativeText = "Original description" }.ToOfficeVisual();
            byte[] payload = visual.GetPlacementBytes();
            using var document = WordDocument.Create();
            var paragraph = document.AddParagraph("Evidence");
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("paragraph", paragraph);
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("visual", visual);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript(preparedVisual
                ? "$visual | WordImage -Paragraph $paragraph -Width 72 -Height 24 -Title 'Placed' -Description 'Service availability' -PassThru"
                : "$paragraph | WordImage -Path $source -Width 72 -Height 24 -Title 'Placed' -Description 'Service availability' -PassThru");
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            var image = Assert.IsType<WordImage>(Assert.Single(result).BaseObject);
            Assert.Equal(OfficeLength.Points(72), image.Width);
            Assert.Equal(OfficeLength.Points(24), image.Height);
            runspace.SessionStateProxy.SetVariable("image", image);
            shell.Commands.Clear();
            shell.AddScript("$image | Set-OfficeWordImage -Width 48 -Height 1cm -Title 'Edited' -PassThru");
            result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Same(image, Assert.Single(result).BaseObject);
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = WordDocument.Load(saved);
            var placed = Assert.Single(reopened.Images);
            Assert.Equal(OfficeLength.Points(48), placed.Width);
            Assert.Equal(OfficeLength.Centimeters(1), placed.Height);
            Assert.Equal("Edited", placed.Title);
            Assert.Equal("Service availability", placed.Description);
            Assert.Equal("Original", visual.Title);
            Assert.Equal("Original description", visual.AlternativeText);
            Assert.Equal(payload, visual.GetPlacementBytes());
        } finally {
            File.Delete(source);
        }
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        foreach (Type type in typeof(AddOfficeWordImageCommand).Assembly.GetTypes()) {
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
