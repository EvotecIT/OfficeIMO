using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Confluence;
using PSWriteOffice.Cmdlets.Confluence;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ConfluenceCmdletTests
{
    [Fact]
    public void PublishPlanOnlyBuildsAdfRequestWithoutSessionOrNetwork()
    {
        using var runspace = CreateRunspace(
            "Publish-OfficeConfluencePage",
            typeof(PublishOfficeConfluencePageCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell
            .AddCommand("Publish-OfficeConfluencePage")
            .AddParameter("SpaceId", "42")
            .AddParameter("Title", "Daily status")
            .AddParameter("Content", "# Ready")
            .AddParameter("PlanOnly");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var plan = Assert.IsType<ConfluencePageWritePlan>(Assert.Single(output).BaseObject);
        Assert.Equal("POST", plan.Method);
        Assert.Contains("atlas_doc_format", plan.Payload);
        Assert.Contains("Daily status", plan.Payload);
    }

    [Fact]
    public void PublishPipelineAggregatesContentIntoOneWritePlan()
    {
        using var runspace = CreateRunspace(
            "Publish-OfficeConfluencePage",
            typeof(PublishOfficeConfluencePageCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell.AddScript("'# First','Second line' | Publish-OfficeConfluencePage -SpaceId 42 -Title 'Pipeline report' -PlanOnly");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var plan = Assert.IsType<ConfluencePageWritePlan>(Assert.Single(output).BaseObject);
        Assert.Contains("First", plan.Payload);
        Assert.Contains("Second line", plan.Payload);
    }

    [Fact]
    public void BearerSessionRequiresAndUsesOAuthCloudId()
    {
        using var runspace = CreateRunspace(
            "New-OfficeConfluenceSession",
            typeof(NewOfficeConfluenceSessionCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        var token = new SecureString();
        foreach (char character in "token") token.AppendChar(character);
        token.MakeReadOnly();
        powerShell
            .AddCommand("New-OfficeConfluenceSession")
            .AddParameter("SiteUri", new Uri("https://example.atlassian.net/"))
            .AddParameter("AccessToken", token)
            .AddParameter("CloudId", "cloud-123");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var session = Assert.IsType<ConfluenceSession>(Assert.Single(output).BaseObject);
        Assert.Equal("https://api.atlassian.com/ex/confluence/cloud-123/", session.ApiBaseUri.AbsoluteUri);
    }

    [Fact]
    public void RemovePlanOnlyBuildsDeleteWithoutSessionOrNetwork()
    {
        using var runspace = CreateRunspace(
            "Remove-OfficeConfluencePage",
            typeof(RemoveOfficeConfluencePageCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell
            .AddCommand("Remove-OfficeConfluencePage")
            .AddParameter("PageId", "123")
            .AddParameter("Purge")
            .AddParameter("PlanOnly");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var plan = Assert.IsType<ConfluencePageWritePlan>(Assert.Single(output).BaseObject);
        Assert.Equal("DELETE", plan.Method);
        Assert.Equal("/wiki/api/v2/pages/123?purge=true", plan.RelativeUri);
    }

    [Fact]
    public void AttachmentDownloadForceStreamsAndAtomicallyReplacesDestination()
    {
        string directory = Path.Combine(Path.GetTempPath(), "PSWriteOffice.Confluence." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "report.txt");
        File.WriteAllText(path, "old");
        try
        {
            using var httpClient = new HttpClient(new StaticResponseHandler(Encoding.UTF8.GetBytes("new-streamed")));
            var session = new ConfluenceSession(
                new ConfluenceBearerCredentialSource("token"),
                new ConfluenceSessionOptions {
                    SiteUri = new Uri("https://example.atlassian.net/"),
                    HttpClient = httpClient
                });
            using var runspace = CreateRunspace(
                "Get-OfficeConfluenceAttachment",
                typeof(GetOfficeConfluenceAttachmentCommand));
            using var powerShell = PowerShell.Create();
            powerShell.Runspace = runspace;
            powerShell
                .AddCommand("Get-OfficeConfluenceAttachment")
                .AddParameter("Session", session)
                .AddParameter("PageId", "123")
                .AddParameter("AttachmentId", "a1")
                .AddParameter("OutFile", path)
                .AddParameter("Force");

            var output = powerShell.Invoke();

            Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
            Assert.IsType<FileInfo>(Assert.Single(output).BaseObject);
            Assert.Equal("new-streamed", File.ReadAllText(path));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void AttachmentDownloadDeletesTemporaryFileWhenStreamingFails()
    {
        string directory = Path.Combine(Path.GetTempPath(), "PSWriteOffice.Confluence." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "report.txt");
        try
        {
            using var httpClient = new HttpClient(new FailingResponseHandler());
            var session = new ConfluenceSession(
                new ConfluenceBearerCredentialSource("token"),
                new ConfluenceSessionOptions {
                    SiteUri = new Uri("https://example.atlassian.net/"),
                    HttpClient = httpClient
                });
            using var runspace = CreateRunspace(
                "Get-OfficeConfluenceAttachment",
                typeof(GetOfficeConfluenceAttachmentCommand));
            using var powerShell = PowerShell.Create();
            powerShell.Runspace = runspace;
            powerShell
                .AddCommand("Get-OfficeConfluenceAttachment")
                .AddParameter("Session", session)
                .AddParameter("PageId", "123")
                .AddParameter("AttachmentId", "a1")
                .AddParameter("OutFile", path);

            var exception = Assert.Throws<CmdletInvocationException>(() => powerShell.Invoke());

            Assert.IsType<HttpRequestException>(exception.InnerException);
            Assert.False(File.Exists(path));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void AzureTableExampleDefersConfluenceTypeResolutionUntilAfterImports()
    {
        string root = FindRepositoryRoot();
        string example = File.ReadAllText(Path.Combine(root, "Examples", "Confluence", "Example-ConfluenceAzureTableReport.ps1"));

        int imports = example.IndexOf("Import-Module PSWriteOffice", StringComparison.Ordinal);
        int runtimeTypeCheck = example.IndexOf("-isnot [OfficeIMO.Confluence.ConfluenceSession]", StringComparison.Ordinal);

        Assert.DoesNotContain("[OfficeIMO.Confluence.ConfluenceSession] $ConfluenceSession", example);
        Assert.True(imports >= 0 && runtimeTypeCheck > imports);
    }

    [Fact]
    public void ManagedSectionCmdletPreservesUnmanagedContent()
    {
        using var runspace = CreateRunspace(
            "Set-OfficeConfluenceManagedSection",
            typeof(SetOfficeConfluenceManagedSectionCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell
            .AddCommand("Set-OfficeConfluenceManagedSection")
            .AddParameter("ExistingBody", "<p>owner</p>")
            .AddParameter("SectionId", "daily")
            .AddParameter("Replacement", "<p>generated</p>")
            .AddParameter("AppendIfMissing");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var result = Assert.IsType<ConfluenceManagedSectionResult>(Assert.Single(output).BaseObject);
        Assert.True(result.WasCreated);
        Assert.Contains("<p>owner</p>", result.UpdatedBody);
        Assert.Contains("<p>generated</p>", result.UpdatedBody);
        Assert.NotEqual(result.OriginalSha256, result.UpdatedSha256);
    }

    [Fact]
    public void LiveWriteCmdletsDeclareShouldProcess()
    {
        AssertShouldProcess(typeof(PublishOfficeConfluencePageCommand));
        AssertShouldProcess(typeof(SendOfficeConfluenceAttachmentCommand));
        AssertShouldProcess(typeof(RemoveOfficeConfluencePageCommand));
    }

    private static Runspace CreateRunspace(string name, Type cmdletType)
    {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry(name, cmdletType, helpFileName: null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static void AssertShouldProcess(Type cmdletType)
    {
        var attribute = Assert.Single(cmdletType.GetCustomAttributes(typeof(CmdletAttribute), inherit: false).Cast<CmdletAttribute>());
        Assert.True(attribute.SupportsShouldProcess);
    }

    private static string FindRepositoryRoot([System.Runtime.CompilerServices.CallerFilePath] string sourceFile = "")
    {
        foreach (string start in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory, Path.GetDirectoryName(sourceFile)! })
        {
            var directory = new DirectoryInfo(start);
            while (directory != null)
            {
                if (Directory.Exists(Path.Combine(directory.FullName, "Examples")) &&
                    Directory.Exists(Path.Combine(directory.FullName, "Sources"))) return directory.FullName;
                directory = directory.Parent;
            }
        }
        throw new DirectoryNotFoundException("PSWriteOffice repository root was not found.");
    }

    private sealed class StaticResponseHandler : HttpMessageHandler
    {
        private readonly byte[] _content;
        internal StaticResponseHandler(byte[] content) => _content = content;

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(_content) });
    }

    private sealed class FailingResponseHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new FailingContent() });
    }

    private sealed class FailingContent : HttpContent
    {
        protected override async Task SerializeToStreamAsync(Stream stream, TransportContext? context)
        {
            await stream.WriteAsync(Encoding.UTF8.GetBytes("partial"));
            throw new IOException("Simulated streaming failure.");
        }

        protected override bool TryComputeLength(out long length)
        {
            length = -1;
            return false;
        }
    }
}
