using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Auth.GoogleApis;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.GoogleWorkspace.Sync;

var options = new GoogleWorkspaceSessionOptions {
    ExpectedAccount = "package-smoke@example.invalid",
    OperationReceiptSink = _ => { },
};
options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
    options.ExpectedAccount!,
    new[] { GoogleWorkspaceScopeCatalog.DriveFile },
    context.Target,
    "package-smoke-revision",
    options.MaxRetryCount,
    options.MaxRetryElapsedTime,
    options.RateLimitPolicy,
    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);

var session = new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource("package-smoke-token"), options);
using var drive = new GoogleDriveClient(session);
var item = new GoogleWorkspaceSyncItem("item-1", GoogleWorkspaceSyncItemKind.SourceChange,
    "document", "package smoke", "drive:file-1", "version:1", googleFileId: "file-1");
var policy = new GoogleWorkspaceOperationPolicy(options.ExpectedAccount!,
    new[] { GoogleWorkspaceScopeCatalog.DriveFile }, "drive:file-1", "version:1",
    options.MaxRetryCount, options.MaxRetryElapsedTime, options.RateLimitPolicy,
    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
GoogleWorkspaceSyncPlan plan = GoogleWorkspaceSyncPlan.Create(new[] { item }, policy);

if (!plan.CanApply || typeof(GoogleDriveDownloadCheckpoint).GetProperty(nameof(GoogleDriveDownloadCheckpoint.ChunkSize)) == null ||
    typeof(GoogleDriveClient).GetMethod(nameof(GoogleDriveClient.DownloadToFileAsync)) == null ||
    typeof(GoogleDriveClient).GetMethod(nameof(GoogleDriveClient.UploadResumableStreamAsync)) == null ||
    typeof(GoogleApisCredentialSource).Assembly == typeof(GoogleWorkspaceSession).Assembly) {
    throw new InvalidOperationException("The packed Google Workspace contracts are incomplete or the optional adapter boundary collapsed.");
}

Console.WriteLine($"OfficeIMO Google Workspace package-family smoke passed on {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}.");
