using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Auth.GoogleApis;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.GoogleWorkspace.Sync;
using System.Net;
using System.Net.Http;
using System.Text;

var receipts = new List<GoogleWorkspaceOperationReceipt>();
var options = new GoogleWorkspaceSessionOptions {
    ExpectedAccount = "package-smoke@example.invalid",
    HttpClient = new HttpClient(new PackageSmokeHandler()),
    OperationReceiptSink = receipts.Add,
};
options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
    options.ExpectedAccount!,
    context.RequiredScopes,
    context.Target,
    context.RevisionPreconditionKind switch {
        GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate => GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
        GoogleWorkspaceRevisionPreconditionKind.PayloadRevision => context.AdapterExpectedRevision!,
        GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState => context.AdapterExpectedRevision!,
        GoogleWorkspaceRevisionPreconditionKind.Unavailable => GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision("package smoke operation"),
        _ => "\"package-smoke-etag\"",
    },
    context.MaxRetryCount,
    context.MaxRetryElapsedTime,
    context.RateLimitPolicy,
    context.RevisionPreconditionKind == GoogleWorkspaceRevisionPreconditionKind.Unavailable
        ? GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss
        : GoogleWorkspaceDataLossDecision.RejectPotentialLoss,
    context.RevisionPreconditionKind == GoogleWorkspaceRevisionPreconditionKind.Unavailable
        ? "package smoke operation without a conditional revision"
        : null);

var session = new GoogleWorkspaceSession(new DelegateGoogleWorkspaceCredentialSource((scopes, _) =>
    Task.FromResult(GoogleWorkspaceAccessToken.FromVerifiedCredential(
        "package-smoke-token", DateTimeOffset.UtcNow.AddMinutes(30),
        new GoogleWorkspaceCredentialBinding(options.ExpectedAccount!, scopes)))), options);
_ = await session.AcquireAccessTokenAsync(new[] { GoogleWorkspaceScopeCatalog.DriveFile });
using var drive = new GoogleDriveClient(session);
using (var transport = new GoogleWorkspaceHttpTransport(session)) {
    _ = await transport.SendJsonAsync<object>(
        "package-smoke-token",
        HttpMethod.Post,
        "https://www.googleapis.com/drive/v3/files",
        new { name = "package-smoke" },
        GoogleWorkspaceRequestSafety.NonIdempotent,
        "Google Drive API",
        new TranslationReport(),
        mutationKind: GoogleWorkspaceMutationKind.Create,
        requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile });
}
var item = new GoogleWorkspaceSyncItem("item-1", GoogleWorkspaceSyncItemKind.SourceChange,
    "document", "package smoke", "drive:file-1", "version:1", googleFileId: "file-1");
var policy = new GoogleWorkspaceOperationPolicy(options.ExpectedAccount!,
    new[] { GoogleWorkspaceScopeCatalog.DriveFile }, "drive:file-1", "version:1",
    options.MaxRetryCount, options.MaxRetryElapsedTime, options.RateLimitPolicy,
    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
GoogleWorkspaceSyncPlan plan = GoogleWorkspaceSyncPlan.Create(new[] { item }, policy);

if (!plan.CanApply || receipts.Count != 1
    || receipts[0].RevisionPreconditionKind != GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate
    || typeof(GoogleDriveDownloadCheckpoint).GetProperty(nameof(GoogleDriveDownloadCheckpoint.ChunkSize)) == null ||
    typeof(GoogleDriveClient).GetMethod(nameof(GoogleDriveClient.DownloadToFileAsync)) == null ||
    typeof(GoogleDriveClient).GetMethod(nameof(GoogleDriveClient.UploadResumableStreamAsync)) == null ||
    typeof(GoogleApisCredentialSource).Assembly == typeof(GoogleWorkspaceSession).Assembly) {
    throw new InvalidOperationException("The packed Google Workspace contracts are incomplete or the optional adapter boundary collapsed.");
}

Console.WriteLine($"OfficeIMO Google Workspace package-family smoke passed on {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}.");

sealed class PackageSmokeHandler : HttpMessageHandler {
    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken) =>
        Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
            Content = new StringContent("{}", Encoding.UTF8, "application/json"),
        });
}
