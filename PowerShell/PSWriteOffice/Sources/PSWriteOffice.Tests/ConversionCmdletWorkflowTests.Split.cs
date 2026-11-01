using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData("-PageCount 2", "1,2|3")]
    [InlineData("-Page '3,1','2'", "3,1|2")]
    [InlineData("-ByBookmark", "1,2|3")]
    [InlineData("-ByBookmark -Bookmark 'Chapter Two'", "3")]
    public void SplitDocumentSavesEachSelectedGroupWithItsOwnPages(string selection, string expectedPages) {
        string root = Directory.CreateTempSubdirectory("officeimo-split-selection-").FullName;
        try {
            string source = Path.Combine(root, "Manual [final].pdf"), destination = Path.Combine(root, "parts");
            WriteSplitFixture(source);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Split-OfficeDocument -LiteralPath $source -OutputDirectory $destination -IndexWidth 2 " + selection + " -ErrorAction Stop");
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            string[] parts = expectedPages.Split('|');
            Assert.Equal(parts.Length, output.Count);
            for (int part = 0; part < parts.Length; part++) {
                var file = Assert.IsType<FileInfo>(output[part].BaseObject);
                Assert.Equal("Manual [final]-" + (part + 1).ToString("D2") + ".pdf", file.Name);
                using var document = PdfDocument.Load(file.FullName);
                string[] pages = parts[part].Split(',');
                var actual = document.Read().Pages;
                Assert.Equal(pages.Length, actual.Count);
                for (int index = 0; index < pages.Length; index++)
                    {
                    using var selected = document.Pages.Extract(PdfPageSelection.Parse((index + 1).ToString()));
                    Assert.Contains("Page " + pages[index] + " evidence", selected.Read().Text, StringComparison.Ordinal);
                }
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SplitDocumentFileInfoPipelineRetainsLiteralNamesAndContinuesAfterRejectedInputs() {
        string root = Directory.CreateTempSubdirectory("officeimo-split-pipeline-").FullName;
        try {
            string source = Path.Combine(root, "scan[1].pdf"), destination = Path.Combine(root, "parts");
            string unsupported = Path.Combine(root, "unknown.bin"), missing = Path.Combine(root, "missing.pdf");
            WriteSplitFixture(source); File.WriteAllText(unsupported, "unsupported");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            runspace.SessionStateProxy.SetVariable("unsupported", unsupported);
            runspace.SessionStateProxy.SetVariable("missing", missing);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                Split-OfficeDocument -LiteralPath $missing,$unsupported -OutputDirectory $destination -ErrorAction Continue
                Get-Item -LiteralPath $source | Split-OfficeDocument -OutputDirectory $destination -PageCount 3 -ErrorAction Stop
                """);
            var output = shell.Invoke();
            Assert.Equal(2, shell.Streams.Error.Count);
            Assert.Contains(shell.Streams.Error, error => error.FullyQualifiedErrorId.StartsWith("input-unreadable", StringComparison.Ordinal));
            Assert.Contains(shell.Streams.Error, error => error.FullyQualifiedErrorId.StartsWith("split-unsupported", StringComparison.Ordinal));
            var file = Assert.IsType<FileInfo>(Assert.Single(output).BaseObject);
            Assert.Equal("scan[1]-1.pdf", file.Name);
            using var document = PdfDocument.Load(file.FullName);
            Assert.Equal(3, document.Inspect().PageCount);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("-WhatIf", false)]
    [InlineData("-Page invalid", true)]
    [InlineData("-ByBookmark:$false", true)]
    [InlineData("-Prefix '../escape'", true)]
    public void SplitDocumentValidatesBeforeEffectsAndWhatIfDoesNotParseThePdf(string arguments, bool fails) {
        string root = Directory.CreateTempSubdirectory("officeimo-split-admission-").FullName;
        try {
            string source = Path.Combine(root, "not-really.pdf"), destination = Path.Combine(root, "parts");
            File.WriteAllText(source, "not a PDF");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Split-OfficeDocument -LiteralPath $source -OutputDirectory $destination " + arguments + " -ErrorAction Stop");
            Assert.Empty(shell.Invoke());
            if (fails) {
                string message = Assert.Single(shell.Streams.Error).Exception.Message;
                string expected = arguments.StartsWith("-ByBookmark", StringComparison.Ordinal) ? "ByBookmark must be enabled" :
                    arguments.StartsWith("-Prefix", StringComparison.Ordinal) ? "Prefix must be a portable" : "Page number";
                Assert.Contains(expected, message, StringComparison.OrdinalIgnoreCase);
            } else Assert.Empty(shell.Streams.Error);
            Assert.Equal("not a PDF", File.ReadAllText(source));
            Assert.False(Directory.Exists(destination));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SplitDocumentNoClobberResolvesSharedBasenamesAtomicallyAndPreservesExistingParts() {
        string root = Directory.CreateTempSubdirectory("officeimo-split-collision-").FullName;
        try {
            string first = Path.Combine(root, "first", "scan.pdf"), second = Path.Combine(root, "second", "scan.pdf");
            Directory.CreateDirectory(Path.GetDirectoryName(first)!); Directory.CreateDirectory(Path.GetDirectoryName(second)!);
            WriteSplitFixture(first); WriteSplitFixture(second);
            string destination = Path.Combine(root, "parts");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("sources", new[] { first, second });
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Split-OfficeDocument -LiteralPath $sources -OutputDirectory $destination -PageCount 3 -NoClobber -ErrorAction Continue");
            var output = shell.Invoke();
            var file = Assert.IsType<FileInfo>(Assert.Single(output).BaseObject);
            Assert.Single(shell.Streams.Error);
            using (var document = PdfDocument.Load(file.FullName)) Assert.Equal(3, document.Inspect().PageCount);
            byte[] retained = File.ReadAllBytes(file.FullName);
            shell.Commands.Clear(); shell.Streams.Error.Clear();
            shell.AddScript("Split-OfficeDocument -LiteralPath $sources -OutputDirectory $destination -PageCount 3 -NoClobber -ErrorAction Continue");
            Assert.Empty(shell.Invoke());
            Assert.Equal(2, shell.Streams.Error.Count);
            Assert.Equal(retained, File.ReadAllBytes(file.FullName));
            Assert.Single(Directory.GetFiles(destination));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SplitDocumentCannotOverwriteAnotherInputEvenWhenReplacementIsAllowed() {
        string root = Directory.CreateTempSubdirectory("officeimo-split-source-protection-").FullName;
        try {
            string source = Path.Combine(root, "scan.pdf"), next = Path.Combine(root, "scan-1.pdf");
            WriteSplitFixture(source); WriteSplitFixture(next);
            byte[] first = File.ReadAllBytes(source), second = File.ReadAllBytes(next);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("sources", new[] { source, next });
            runspace.SessionStateProxy.SetVariable("destination", root);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Split-OfficeDocument -LiteralPath $sources -OutputDirectory $destination -PageCount 3 -ErrorAction Continue");
            Assert.Single(shell.Invoke());
            Assert.Contains("overwrite an input", Assert.Single(shell.Streams.Error).Exception.Message, StringComparison.Ordinal);
            Assert.Equal(first, File.ReadAllBytes(source)); Assert.Equal(second, File.ReadAllBytes(next));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("open", true)]
    [InlineData("wrong", false)]
    public void SplitDocumentAuthenticatesBeforeUsingPermissionOverridesAndKeepsEncryption(string password, bool succeeds) {
        string root = Directory.CreateTempSubdirectory("officeimo-split-password-").FullName;
        try {
            string source = Path.Combine(root, "protected.pdf"), destination = Path.Combine(root, "parts");
            using (var document = PdfDocument.Create(new PdfCreateOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
            }))) {
                document.Content.Paragraph(p => p.Text("Authenticated split evidence"));
                document.Save(source);
            }
            var settings = new PdfLoadOptions { PermissionPolicy = PdfPermissionPolicy.Enforce };
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            runspace.SessionStateProxy.SetVariable("password", password);
            runspace.SessionStateProxy.SetVariable("settings", settings);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Split-OfficeDocument -LiteralPath $source -OutputDirectory $destination -Password (ConvertTo-SecureString $password -AsPlainText -Force) -Options @($settings, @{ PermissionPolicy = 'IgnoreRestrictions' }) -ErrorAction Continue");
            var output = shell.Invoke();
            if (succeeds) {
                Assert.Empty(shell.Streams.Error);
                var file = Assert.IsType<FileInfo>(Assert.Single(output).BaseObject);
                using var part = PdfDocument.Load(file.FullName, new PdfLoadOptions { Password = "open", PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions });
                Assert.True(part.Inspect().Security.HasEncryption);
                Assert.Contains("Authenticated split evidence", part.Read().Text, StringComparison.Ordinal);
                using var unauthenticated = PdfDocument.Load(file.FullName);
                Assert.ThrowsAny<Exception>(() => unauthenticated.Read());
            } else {
                Assert.Empty(output); Assert.Single(shell.Streams.Error);
                Assert.False(Directory.Exists(destination));
            }
            Assert.Equal(PdfPermissionPolicy.Enforce, settings.PermissionPolicy);
        } finally { Directory.Delete(root, true); }
    }

    private static void WriteSplitFixture(string path) {
        using var document = PdfDocument.Create(new PdfCreateOptions { CreateOutlineFromHeadings = true });
        document.Content.H1("Chapter One").Paragraph(p => p.Text("Page 1 evidence"))
            .PageBreak().H2("Nested section").Paragraph(p => p.Text("Page 2 evidence"))
            .PageBreak().H1("Chapter Two").Paragraph(p => p.Text("Page 3 evidence"));
        document.Save(path);
    }
}
