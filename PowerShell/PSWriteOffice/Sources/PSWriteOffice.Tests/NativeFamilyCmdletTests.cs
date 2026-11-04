using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Epub;
using PSWriteOffice.Cmdlets.Access;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class NativeFamilyCmdletTests {
    [Theory]
    [InlineData("Accdb", ".accdb")]
    [InlineData("Mdb", ".mdb")]
    public void AccessObjectsCreateNativeDatabaseAndStreamTypedRows(string format, string extension) {
        InDirectory(directory => {
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("outputPath", Path.Combine(directory, "contacts" + extension));
            runspace.SessionStateProxy.SetVariable("databaseFormat", format);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $database = New-OfficeAccess -Format $databaseFormat -Title Contacts
                try {
                    $table = $database | Add-OfficeAccessTable -Name Contacts -PassThru
                    $table | Add-OfficeAccessColumn -Name Id -Type AutoNumber
                    $table | Add-OfficeAccessColumn -Name DisplayName -MaxLength 120
                    $table | Add-OfficeAccessColumn -Name Enabled -Type Boolean
                    $table | Add-OfficeAccessColumn -Name Count -Type Int32
                    @(
                        [pscustomobject]@{ DisplayName = 'Ada'; Enabled = 'False'; Count = '2' }
                        @{ DisplayName = 'Grace'; Enabled = 'True'; Count = '3' }
                    ) | Add-OfficeAccessRow -Table $table
                    $database | Save-OfficeAccess -Path $outputPath -WhatIf
                    if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf wrote a database' }
                    $database | Save-OfficeAccess -Path $outputPath
                } finally {
                    $database | Close-OfficeAccess
                }
                $saved = Get-Item -LiteralPath $outputPath | Get-OfficeAccess
                try {
                    $saved | Get-OfficeAccessRow -TableName Contacts
                } finally {
                    $saved | Close-OfficeAccess
                }
                """);
            var rows = shell.Invoke();
            AssertNoErrors(shell);
            Assert.Equal(2, rows.Count);
            Assert.Equal("Ada", rows[0].Properties["DisplayName"].Value);
            Assert.Equal(false, rows[0].Properties["Enabled"].Value);
            Assert.Equal(2, rows[0].Properties["Count"].Value);
            Assert.Equal(true, rows[1].Properties["Enabled"].Value);
            Assert.Equal(1, rows[0].Properties["Id"].Value);
            Assert.Equal(2, rows[1].Properties["Id"].Value);
        });
    }

    [Fact]
    public void EpubChaptersAndMetadataSurviveSaveAndPackageReadback() {
        InDirectory(directory => {
            using var runspace = CreateRunspace();
            string path = Path.Combine(directory, "handbook.epub");
            runspace.SessionStateProxy.SetVariable("outputPath", path);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $book = New-OfficeEpub -Path $outputPath -Title Operations -Language en -Author Example
                if (Test-Path -LiteralPath $outputPath) { throw 'Live creation wrote an EPUB' }
                $book | Add-OfficeEpubChapter -Title Opening -Text 'Safety & care'
                $book | Add-OfficeEpubChapter -Title Procedures -Markdown '# Procedure A'
                $book | Set-OfficeEpubMetadata -Title 'Operations handbook'
                $book | Save-OfficeEpub -WhatIf
                if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf wrote an EPUB' }
                $evidence = $book | Save-OfficeEpub -PassThru
                Get-Item -LiteralPath $outputPath | Open-OfficeEpub | Get-OfficeEpubChapter
                """);
            var chapters = shell.Invoke();
            AssertNoErrors(shell);
            Assert.Equal(2, chapters.Count);
            Assert.Contains(chapters, chapter => Assert.IsType<EpubChapter>(chapter.BaseObject).Title == "Opening");
            var saved = EpubPublication.Load(path);
            Assert.Equal("Operations handbook", saved.Title);
            Assert.Equal("Example", saved.Creator);
            Assert.Equal(2, saved.Spine.Count);
            Assert.False(Assert.IsType<EpubWriteReport>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("evidence")).BaseObject).HasLoss);
        });
    }

    private static Runspace CreateRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(NewOfficeAccessCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static void AssertNoErrors(PowerShell shell) =>
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));

    private static void InDirectory(Action<string> action) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-families-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            action(directory);
        } finally {
            Directory.Delete(directory, true);
        }
    }
}
