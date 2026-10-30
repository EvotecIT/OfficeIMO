using System.Collections;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Provenance;
using PSWriteOffice.Cmdlets.Security;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ProvenanceCmdletTests {
    [Fact]
    public void PathsExpandWildcardsAndPreserveLiteralFileInfoNames() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-provenance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string first = Path.Combine(directory, "press[1].txt");
            string second = Path.Combine(directory, "press[2].txt");
            File.WriteAllText(first, "First source");
            File.WriteAllText(second, "Second source");
            using var runspace = CreateRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Get-OfficeDocumentProvenance").AddParameter("Path", Path.Combine(directory, "*.txt"));
            Assert.Equal(2, shell.Invoke().Count);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            shell.Commands.Clear();
            shell.AddCommand("Get-OfficeDocumentProvenance").AddParameter("LiteralPath", new[] { first, second });
            Assert.Equal(2, shell.Invoke().Count);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            shell.Commands.Clear();
            shell.AddCommand("Get-OfficeDocumentProvenance");
            Assert.Equal(2, shell.Invoke(new[] { new FileInfo(first), new FileInfo(second) }).Count);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal("First source", File.ReadAllText(first));
            Assert.Equal("Second source", File.ReadAllText(second));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void TypedAndNestedHashtablePoliciesControlActualTextInspection() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-provenance-" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "Hidden\u200b text");
        try {
            using var runspace = CreateRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-OfficeDocumentProvenanceOptions").AddParameter("NoTextIntegrity")
                .AddParameter("MaximumAssetBytes", 1024L).AddParameter("MaximumManifestBytes", 1024L)
                .AddParameter("VerificationTimeoutSeconds", 60);
            var policy = Assert.IsType<OfficeProvenanceAssessmentOptions>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(1024L, policy.Structural.MaxAssetBytes);
            Assert.Equal(TimeSpan.FromSeconds(60), policy.Verification.Timeout);
            Assert.False(policy.Verification.AllowNetworkAccess);
            shell.Commands.Clear();
            shell.AddCommand("Get-OfficeDocumentProvenance").AddParameter("LiteralPath", path).AddParameter("Options", policy);
            var disabledResults = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            var disabled = Assert.IsType<OfficeProvenanceAssessmentReport>(Assert.Single(disabledResults).BaseObject);
            Assert.Equal(OfficeProvenanceCheckStatus.Disabled, disabled.TextIntegrityStatus);
            Assert.Null(disabled.Verification);
            shell.Commands.Clear();
            shell.AddCommand("Get-OfficeDocumentProvenance").AddParameter("LiteralPath", path).AddParameter("Options", new Hashtable {
                ["Structural"] = new Hashtable { ["MaxAssetBytes"] = 1024L, ["MaxManifestBytes"] = 1024L },
                ["TextIntegrity"] = new Hashtable { ["MaxFindings"] = 1 }
            });
            var inspectedResults = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            var inspected = Assert.IsType<OfficeProvenanceAssessmentReport>(Assert.Single(inspectedResults).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(OfficeProvenanceCheckStatus.Completed, inspected.TextIntegrityStatus);
            Assert.NotNull(inspected.TextIntegrity);
            Assert.NotEmpty(inspected.TextIntegrity!.Findings);
            shell.Commands.Clear();
            shell.AddCommand("Get-OfficeDocumentProvenance").AddParameter("LiteralPath", path)
                .AddParameter("Options", new Hashtable { ["Structural"] = new Hashtable { ["UnknownLimit"] = 1 } });
            Assert.ThrowsAny<ParameterBindingException>(() => shell.Invoke());
            Assert.Equal("Hidden\u200b text", File.ReadAllText(path));
        } finally { File.Delete(path); }
    }

    private static Runspace CreateRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Get-OfficeDocumentProvenance", typeof(GetOfficeDocumentProvenanceCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeDocumentProvenanceOptions", typeof(NewOfficeDocumentProvenanceOptionsCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
