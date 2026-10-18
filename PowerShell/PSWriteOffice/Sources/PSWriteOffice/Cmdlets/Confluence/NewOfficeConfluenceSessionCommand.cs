using System;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security;
using OfficeIMO.Confluence;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Creates an in-memory Confluence Cloud session.</summary>
/// <example>
/// <summary>Create a session using an Atlassian email and API token.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>$session = New-OfficeConfluenceSession -SiteUri 'https://example.atlassian.net/' -Credential (Get-Credential)</code>
/// <para>Use the Atlassian email as the user name and the API token as the password.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeConfluenceSession", DefaultParameterSetName = ParameterSetBasic)]
[OutputType(typeof(ConfluenceSession))]
public sealed class NewOfficeConfluenceSessionCommand : PSCmdlet
{
    private const string ParameterSetBasic = "Basic";
    private const string ParameterSetBearer = "Bearer";

    /// <summary>HTTPS root URI of the Confluence Cloud site.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public Uri SiteUri { get; set; } = null!;

    /// <summary>Atlassian email and API token credential.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetBasic)]
    public PSCredential Credential { get; set; } = null!;

    /// <summary>OAuth access token stored as a secure string.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetBearer)]
    public SecureString AccessToken { get; set; } = null!;

    /// <summary>Atlassian Cloud identifier required for OAuth bearer-token routing.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetBearer)]
    [ValidateNotNullOrEmpty]
    public string CloudId { get; set; } = string.Empty;

    /// <summary>Application name sent with requests.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string ApplicationName { get; set; } = "PSWriteOffice";

    /// <summary>Per-request timeout in seconds.</summary>
    [Parameter]
    [ValidateRange(1, 3600)]
    public int RequestTimeoutSeconds { get; set; } = 100;

    /// <summary>Maximum retry count for safe read requests. Writes are never retried automatically.</summary>
    [Parameter]
    [ValidateRange(0, 10)]
    public int MaxRetryCount { get; set; } = 3;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        IConfluenceCredentialSource credentialSource = ParameterSetName == ParameterSetBearer
            ? new ConfluenceBearerCredentialSource(ReadSecureString(AccessToken))
            : new ConfluenceBasicCredentialSource(Credential.UserName, Credential.GetNetworkCredential().Password);
        WriteObject(new ConfluenceSession(credentialSource, new ConfluenceSessionOptions
        {
            SiteUri = SiteUri,
            CloudId = ParameterSetName == ParameterSetBearer ? CloudId : null,
            ApplicationName = ApplicationName,
            RequestTimeout = TimeSpan.FromSeconds(RequestTimeoutSeconds),
            MaxRetryCount = MaxRetryCount
        }));
    }

    private static string ReadSecureString(SecureString value)
    {
        IntPtr pointer = IntPtr.Zero;
        try
        {
            pointer = Marshal.SecureStringToBSTR(value);
            return Marshal.PtrToStringBSTR(pointer);
        }
        finally
        {
            if (pointer != IntPtr.Zero)
            {
                Marshal.ZeroFreeBSTR(pointer);
            }
        }
    }
}
