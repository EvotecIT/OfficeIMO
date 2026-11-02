using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Epub;
using PSWriteOffice.Cmdlets.Epub;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class EpubLifecycleWorkflowTests {
    [Fact]
    public void ChapterAndResourceBlocksUseTheNativeContextAndPublishOnlyAfterSuccess() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-epub-block-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Guide[final].epub"), css = Path.Combine(root, "book.css");
            File.WriteAllText(css, "body { color: #123456; }");
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(NewOfficeEpubCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("bookPath", path);
            runspace.SessionStateProxy.SetVariable("cssPath", css);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                New-OfficeEpub -Path $bookPath -Title Guide {
                    EpubChapter -Title Opening -Text 'Hello & welcome'
                    EpubResource -Path $cssPath -Id style -MediaType text/css -ContainerPath EPUB/styles/book.css
                }
                Open-OfficeEpub -LiteralPath $bookPath { EpubChapter -Title Review -Text Approved }
                """);
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            using EpubPublication saved = EpubPublication.Load(path);
            Assert.Equal("Guide", saved.Title);
            Assert.Equal(2, saved.Spine.Count);
            Assert.Contains(saved.Read().Chapters, chapter => chapter.Text.Contains("Hello & welcome"));
            using (var package = ZipFile.OpenRead(path)) {
                using var resource = new StreamReader(package.GetEntry("EPUB/styles/book.css")!.Open());
                Assert.Equal("body { color: #123456; }", resource.ReadToEnd());
            }
            byte[] original = File.ReadAllBytes(path);
            shell.Commands.Clear();
            shell.AddCommand("New-OfficeEpub").AddParameter("Path", path).AddParameter("Title", "Discarded")
                .AddParameter("Force").AddParameter("Content", ScriptBlock.Create("EpubChapter -Title Discarded -Text 'Do not save'; throw 'authoring-failed'"));
            Exception? error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            Assert.Contains("authoring-failed", error?.Message ?? string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(2, Directory.GetFiles(root).Length);
        } finally {
            Directory.Delete(root, true);
        }
    }
}
