using System.Threading.Tasks;
using PSWriteOffice.Services;
using System;
using System.Management.Automation;
using OfficeIMO.GoogleWorkspace;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Google;

namespace PSWriteOffice.Cmdlets.Google;

/// <summary>Creates a native Google Workspace session from a service-account key.</summary>
/// <para>Authenticates the requested scopes through the native JWT credential source. The returned session owns no browser or persistent token store. Dispose its CredentialSource when finished. Service-account PEM import requires PowerShell 7.4 or newer.</para>
/// <example><summary>Create a delegated reporting session.</summary><code>$session = New-OfficeGoogleWorkspaceSession -KeyPath ./Secrets/reporting.json -Subject reports@example.com -Scope Documents,DriveFile
/// try { Export-OfficeWordGoogleDocument -Path ./Report.docx -Session $session } finally { $session.CredentialSource.Dispose() }</code></example>
[Cmdlet(VerbsCommon.New, "OfficeGoogleWorkspaceSession", DefaultParameterSetName = "ServiceAccount")]
[OutputType(typeof(GoogleWorkspaceSession))]
public sealed class NewOfficeGoogleWorkspaceSessionCommand : AsyncPSCmdlet {
    /// <summary>Exact service-account JSON key path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ServiceAccount")]
    public string KeyPath { get; set; } = string.Empty;
    /// <summary>Delegated user; omission acts as the service account itself.</summary>
    [Parameter(ParameterSetName = "ServiceAccount")]
    public string? Subject { get; set; }
    /// <summary>Native scope names or full URLs; default authorizes Docs and Sheets with DriveFile.</summary>
    [Parameter(ParameterSetName = "ServiceAccount"), ArgumentCompleter(typeof(OfficeGoogleWorkspaceScopeCompleter))]
    public string[]? Scope { get; set; }
    /// <summary>Application name sent with requests.</summary>
    [Parameter(ParameterSetName = "ServiceAccount")]
    public string ApplicationName { get; set; } = "PSWriteOffice";
    /// <summary>Per-request timeout.</summary>
    [Parameter(ParameterSetName = "ServiceAccount"), ValidateRange(1, 3600)]
    public int RequestTimeoutSeconds { get; set; } = 100;
    /// <summary>Maximum native retries.</summary>
    [Parameter(ParameterSetName = "ServiceAccount"), ValidateRange(0, 10)]
    public int MaxRetryCount { get; set; } = 3;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        var scopes = GoogleWorkspaceScopes.Resolve(Scope);
        var options = new GoogleWorkspaceSessionOptions {
            ApplicationName = ApplicationName, RequestTimeout = TimeSpan.FromSeconds(RequestTimeoutSeconds), MaxRetryCount = MaxRetryCount,
            SubjectUser = Subject, UseDomainWideDelegation = Subject != null
        };
        var source = GoogleServiceAccountCredentialSource.FromFile(OfficePath.ResolveInput(SessionState, KeyPath, literal: true)[0], options);
        try {
            var session = new GoogleWorkspaceSession(source, options);
            await session.AcquireAccessTokenAsync(scopes, CancelToken).ConfigureAwait(false);
            WriteObject(session);
        }
        catch { source.Dispose(); throw; }
    }
}
