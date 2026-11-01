using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Markdown;
using PSWriteOffice.Cmdlets.Markdown;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class MarkdownLifecycleWorkflowTests {
    [Fact]
    public void LiveCreationDoesNotWriteAndNativeSaveReturnsTheAssociatedFile() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-markdown-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "live.md");
            using var runspace = Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-OfficeMarkdown").AddParameter("Path", path);
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.False(File.Exists(path), "Creating a live document must not publish a file.");
            using var document = Assert.IsType<MarkdownDocument>(Assert.Single(result).BaseObject);
            document.H1("Live report").P("Saved explicitly");
            var options = new MarkdownSaveOptions { OutputLineEnding = "\n" };
            shell.Commands.Clear();
            shell.AddCommand("Save-OfficeMarkdown").AddParameter("Document", document)
                .AddParameter("WriteOptions", options).AddParameter("LineEnding", "CRLF").AddParameter("PassThru");
            FileInfo file = Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(path, file.FullName);
            Assert.Contains("\r\n", File.ReadAllText(path), StringComparison.Ordinal);
            Assert.Equal("\n", options.OutputLineEnding);
            using var reopened = MarkdownDocument.Load(path);
            Assert.Contains("Live report", reopened.ToMarkdown(), StringComparison.Ordinal);
            byte[] saved = File.ReadAllBytes(path);
            document.P("Discard this edit");
            shell.Commands.Clear();
            shell.AddCommand("Close-OfficeMarkdown").AddParameter("Document", document).Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(saved, File.ReadAllBytes(path));
        } finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public void ContentBlocksPublishOnlyAfterSuccessAndWhatIfDoesNotEnterTheBlock() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-markdown-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "existing.md");
            File.WriteAllText(path, "# Keep this source\n");
            byte[] original = File.ReadAllBytes(path);
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("entered", false);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-OfficeMarkdown").AddParameter("Path", path).AddParameter("Force")
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; MarkdownHeading -Text 'Discard'; throw 'markdown-block-failed'"));
            var failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.True((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(folder));

            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            runspace.SessionStateProxy.SetVariable("entered", false);
            shell.AddCommand("New-OfficeMarkdown").AddParameter("Path", path).AddParameter("Force").AddParameter("WhatIf")
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; MarkdownHeading -Text 'Skipped'"));
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));

            shell.Commands.Clear();
            shell.AddCommand("New-OfficeMarkdown").AddParameter("Path", path).AddParameter("Force").AddParameter("PassThru")
                .AddParameter("Content", ScriptBlock.Create("MarkdownHeading -Text 'Published'"));
            Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            using var reopened = MarkdownDocument.Load(path);
            Assert.Contains("# Published", reopened.ToMarkdown(), StringComparison.Ordinal);
            Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public void NativeSaveRejectsSelectedFrontMatterLossUnlessExplicitlyAccepted() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-markdown-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "loss.md");
            File.WriteAllText(path, "# Existing\n");
            byte[] original = File.ReadAllBytes(path);
            using var document = MarkdownDocument.Parse("---\ntitle: Report\n---\n\n# Body");
            Assert.True(document.HasDocumentHeader);
            var options = MarkdownSaveOptions.CreateCommonMarkProfile();
            using var runspace = Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Save-OfficeMarkdown").AddParameter("Document", document).AddParameter("Path", path)
                .AddParameter("WriteOptions", options);
            var failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.Equal(original, File.ReadAllBytes(path));
            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            shell.AddCommand("Save-OfficeMarkdown").AddParameter("Document", document).AddParameter("Path", path)
                .AddParameter("WriteOptions", options).AddParameter("AllowLoss").AddParameter("PassThru");
            Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.DoesNotContain("title:", File.ReadAllText(path), StringComparison.Ordinal);
            Assert.Contains("# Body", File.ReadAllText(path), StringComparison.Ordinal);
            Assert.False(options.AllowLoss);
        } finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public void OpenUsesLiteralFilePipelinesAndWildcardsAndSnapshotsBoundedReaderOptions() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-markdown-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "source[final].md");
            File.WriteAllText(path, "# Source\n\nRead this document.");
            File.WriteAllText(Path.Combine(folder, "other.md"), "# Other");
            byte[] original = File.ReadAllBytes(path);
            using var runspace = Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-OfficeMarkdownReaderOptions").AddParameter("MaxInputCharacters", 256)
                .AddParameter("DisallowFileUrls");
            var reader = Assert.IsType<MarkdownLoadOptions>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.True(reader.DisallowFileUrls);
            reader.InputNormalization.NormalizeZeroWidthSpacingArtifacts = false;
            shell.Commands.Clear();
            shell.AddCommand("Open-OfficeMarkdown").AddParameter("Options", new object[] { reader,
                new System.Collections.Hashtable { ["InputNormalization"] = reader.InputNormalization },
                new System.Collections.Hashtable { ["InputNormalization"] = new System.Collections.Hashtable {
                    ["NormalizeZeroWidthSpacingArtifacts"] = true } } });
            using var loaded = Assert.IsType<MarkdownDocument>(Assert.Single(shell.Invoke(new[] { new FileInfo(path) })).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(path, loaded.FilePath);
            Assert.False(reader.InputNormalization.NormalizeZeroWidthSpacingArtifacts);
            shell.Commands.Clear();
            shell.AddCommand("Open-OfficeMarkdown").AddParameter("Path", Path.Combine(folder, "*.md"));
            var documents = shell.Invoke().Select(item => Assert.IsType<MarkdownDocument>(item.BaseObject)).ToArray();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(2, documents.Length);
            foreach (var document in documents) document.Dispose();

            shell.Commands.Clear();
            shell.AddCommand("Open-OfficeMarkdown").AddParameter("LiteralPath", path)
                .AddParameter("Options", new object[] { reader, new System.Collections.Hashtable { ["MaxInputCharacters"] = 4 } });
            var failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.Equal(256, reader.MaxInputCharacters);
            Assert.Equal(original, File.ReadAllBytes(path));
            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            runspace.SessionStateProxy.SetVariable("entered", false);
            shell.AddCommand("Open-OfficeMarkdown").AddParameter("LiteralPath", path)
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; MarkdownParagraph -Text 'Discard'; throw 'open-markdown-failed'"));
            failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.True((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(2, Directory.GetFiles(folder).Length);

            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            runspace.SessionStateProxy.SetVariable("entered", false);
            shell.AddCommand("Open-OfficeMarkdown").AddParameter("Text", "# In memory")
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true"));
            failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally { Directory.Delete(folder, true); }
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        foreach (Type type in typeof(NewOfficeMarkdownCommand).Assembly.GetTypes()
            .Where(type => type.Namespace == typeof(NewOfficeMarkdownCommand).Namespace && !type.IsAbstract)) {
            var attribute = (CmdletAttribute?)Attribute.GetCustomAttribute(type, typeof(CmdletAttribute));
            if (attribute == null) continue;
            string name = attribute.VerbName + "-" + attribute.NounName;
            state.Commands.Add(new SessionStateCmdletEntry(name, type, null));
            var aliases = (AliasAttribute?)Attribute.GetCustomAttribute(type, typeof(AliasAttribute));
            foreach (string alias in aliases?.AliasNames ?? Array.Empty<string>())
                state.Commands.Add(new SessionStateAliasEntry(alias, name));
        }
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
