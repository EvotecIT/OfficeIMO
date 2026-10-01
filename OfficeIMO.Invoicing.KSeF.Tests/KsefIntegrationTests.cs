using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Invoicing.KSeF;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.KSeF.Tests;

public class KsefIntegrationTests {
    [Fact]
    public async Task CloseWaitsForInFlightSubmissionAndDisposalPreventsSecretReuse() {
        using var harness = new KsefHarness { SubmissionStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously), ReleaseSubmission = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously) };
        using var client = new KsefClient(handler: harness, clock: new KsefClock()); using KsefCredentials credentials = await Authenticate(client, harness);
        using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials);
        Task<KsefSubmission> pending = client.SubmitAsync(credentials, session, InvoiceBytes(), Validator());
        await harness.SubmissionStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Task close = client.CloseOnlineSessionAsync(credentials, session);
        Assert.DoesNotContain(harness.Requests, request => request.EndsWith("/close", StringComparison.Ordinal));
        harness.ReleaseSubmission.TrySetResult(); Assert.NotNull(await pending); await close;
        credentials.Dispose(); await Assert.ThrowsAsync<ObjectDisposedException>(() => client.OpenOnlineSessionAsync(credentials));
        using var disposed = new KsefSecret("fictional-token"); Assert.Equal("[redacted]", disposed.ToString());
        Assert.Equal("{}", System.Text.Json.JsonSerializer.Serialize(disposed)); disposed.Dispose();
        int calls = harness.Requests.Count; await Assert.ThrowsAsync<ObjectDisposedException>(() => client.BeginTokenAuthenticationAsync(new KsefContext(KsefContextKind.Nip, "9999999999"), disposed));
        Assert.Equal(calls, harness.Requests.Count);
    }
    private static Fa3SchemaValidator Validator() => new(Fa3SchemaBundle.LoadDirectory(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "Schemas")));
    private static byte[] InvoiceBytes() => Fa3InvoiceWriter.Write(Fa3InvoiceFixture.Create(), Fa3InvoiceFixture.Options());
    private static async Task<KsefCredentials> Authenticate(KsefClient client, KsefHarness harness) {
        harness.ExpectedToken = "fictional-ksef-token";
        using var secret = new KsefSecret(harness.ExpectedToken); using KsefAuthentication authentication = await client.BeginTokenAuthenticationAsync(new KsefContext(KsefContextKind.Nip, "9999999999"), secret);
        Assert.True((await client.GetAuthenticationStatusAsync(authentication)).IsSuccessful);
        return await client.RedeemAsync(authentication);
    }
    [Fact]
    public async Task ExactValidatedBytesFlowThroughTokenAuthRsaAesStatusAndBoundReceipt() {
        using var harness = new KsefHarness(); using var client = new KsefClient(handler: harness, clock: new KsefClock()); using KsefCredentials credentials = await Authenticate(client, harness);
        await client.RefreshAsync(credentials); using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials);
        byte[] input = InvoiceBytes(), original = (byte[])input.Clone(); harness.OnSubmission = () => Array.Clear(input);
        KsefSubmission submission = await client.SubmitAsync(credentials, session, input, Validator());
        Assert.Equal(original, harness.Plaintext); Assert.Equal(Convert.ToBase64String(SHA256.HashData(original)), submission.InvoiceHash);
        harness.PendingInvoice = true; Assert.False((await client.GetInvoiceStatusAsync(credentials, submission)).IsAccepted);
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.GetInvoiceReceiptAsync(credentials, submission));
        harness.PendingInvoice = false; Assert.True((await client.GetInvoiceStatusAsync(credentials, submission)).IsAccepted);
        KsefReceipt receipt = await client.GetInvoiceReceiptAsync(credentials, submission);
        Assert.True(receipt.IsBound); Assert.True(receipt.RetrievedThroughAuthenticatedApi); Assert.Equal(InvoiceValidationStatus.Passed, receipt.SchemaStatus);
        byte[] exposed = receipt.GetBytes(); exposed[0] ^= 1; Assert.NotEqual(exposed, receipt.GetBytes());
        await client.CloseOnlineSessionAsync(credentials, session);
        Assert.True((await client.GetSessionStatusAsync(credentials, session.ReferenceNumber)).IsSuccessful);
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.SubmitAsync(credentials, session, original, Validator()));
        Assert.DoesNotContain(harness.Requests, request => request.Contains("untrusted", StringComparison.Ordinal));
    }
    [Theory]
    [InlineData("network")]
    [InlineData("cancel")]
    [InlineData("server")]
    [InlineData("malformed")]
    [InlineData("declared-size")]
    [InlineData("stream-size")]
    [InlineData("redirected-response")]
    public async Task UnknownSubmissionOutcomeStopsMutationAndAllowsReadOnlyHashReconciliation(string failure) {
        using var harness = new KsefHarness { Failure = failure }; using var client = new KsefClient(handler: harness, clock: new KsefClock()); using KsefCredentials credentials = await Authenticate(client, harness);
        using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials); byte[] xml = InvoiceBytes();
        KsefMutationAmbiguousException exception = await Assert.ThrowsAsync<KsefMutationAmbiguousException>(() => client.SubmitAsync(credentials, session, xml, Validator()));
        Assert.Equal(session.ReferenceNumber, exception.ReferenceNumber); Assert.Equal(harness.PlaintextHash, exception.InvoiceHash);
        Assert.DoesNotContain("fictional-ksef-token", exception.ToString()); Assert.DoesNotContain("raw-sensitive-server-body", exception.ToString());
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.SubmitAsync(credentials, session, xml, Validator()));
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.CloseOnlineSessionAsync(credentials, session));
        KsefInvoicePage page = await client.GetSessionInvoicesAsync(credentials, session.ReferenceNumber);
        KsefSessionInvoice reconciled = Assert.Single(page.Invoices); Assert.Equal(exception.InvoiceHash, reconciled.Submission.InvoiceHash);
        Assert.True((await client.GetInvoiceStatusAsync(credentials, reconciled.Submission)).IsAccepted);
        Assert.Single(harness.Requests, request => request.EndsWith("/invoices", StringComparison.Ordinal) && request.StartsWith("POST", StringComparison.Ordinal));
    }
    [Theory]
    [InlineData("declared-size")]
    [InlineData("stream-size")]
    [InlineData("redirected-response")]
    public async Task UnknownCloseResponseStopsMutationAndRetainsReadableSessionReference(string failure) {
        using var harness = new KsefHarness { CloseFailure = failure }; using var client = new KsefClient(handler: harness, clock: new KsefClock());
        using KsefCredentials credentials = await Authenticate(client, harness); using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials);
        KsefMutationAmbiguousException error = await Assert.ThrowsAsync<KsefMutationAmbiguousException>(() => client.CloseOnlineSessionAsync(credentials, session));
        Assert.Equal(session.ReferenceNumber, error.ReferenceNumber); Assert.Null(error.InvoiceHash);
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.CloseOnlineSessionAsync(credentials, session));
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.SubmitAsync(credentials, session, InvoiceBytes(), Validator()));
        Assert.True((await client.GetSessionStatusAsync(credentials, error.ReferenceNumber!)).IsSuccessful);
        Assert.Single(harness.Requests, request => request.EndsWith("/close", StringComparison.Ordinal));
    }
    [Fact]
    public async Task InvalidXmlAndPredispatchCancellationDoNotSubmitAndExplicitHttpRejectionDoesNotLeakBody() {
        using var harness = new KsefHarness(); using var client = new KsefClient(handler: harness, clock: new KsefClock()); using KsefCredentials credentials = await Authenticate(client, harness);
        using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials);
        await Assert.ThrowsAsync<InvalidDataException>(() => client.SubmitAsync(credentials, session, Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY secret 'x'>]><x/>"), Validator()));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => client.SubmitAsync(credentials, session, InvoiceBytes(), Validator(), cancelled.Token));
        Assert.DoesNotContain(harness.Requests, request => request.EndsWith("/invoices", StringComparison.Ordinal));
        harness.Failure = "reject"; KsefApiException rejection = await Assert.ThrowsAsync<KsefApiException>(() => client.SubmitAsync(credentials, session, InvoiceBytes(), Validator()));
        Assert.DoesNotContain("raw-sensitive-server-body", rejection.ToString()); Assert.Equal(System.Net.HttpStatusCode.BadRequest, rejection.StatusCode);
        harness.Failure = null; Assert.NotNull(await client.SubmitAsync(credentials, session, InvoiceBytes(), Validator()));
    }
    [Theory]
    [InlineData("malformed")]
    [InlineData("declared-size")]
    [InlineData("stream-size")]
    [InlineData("redirected-response")]
    public async Task OneTimeRedemptionCannotReplayAfterALostTokenResponse(string failure) {
        using var harness = new KsefHarness { ExpectedToken = "fictional-token", RedemptionFailure = failure }; using var client = new KsefClient(handler: harness, clock: new KsefClock());
        using var secret = new KsefSecret(harness.ExpectedToken); using KsefAuthentication authentication = await client.BeginTokenAuthenticationAsync(new KsefContext(KsefContextKind.Nip, "9999999999"), secret);
        harness.RedemptionPending = true; await Assert.ThrowsAsync<InvalidOperationException>(() => client.RedeemAsync(authentication));
        Assert.DoesNotContain("POST auth/token/redeem", harness.Requests); harness.RedemptionPending = false;
        await Assert.ThrowsAsync<KsefMutationAmbiguousException>(() => client.RedeemAsync(authentication));
        await Assert.ThrowsAsync<InvalidOperationException>(() => client.RedeemAsync(authentication));
        Assert.Single(harness.Requests, request => request == "POST auth/token/redeem");
    }
    [Fact]
    public async Task CertificateDigestContextOwnershipAndStatusHashAreEnforced() {
        using var harness = new KsefHarness { InvalidPublicKeyId = true }; var clock = new KsefClock();
        using var client = new KsefClient(handler: harness, clock: clock);
        await Assert.ThrowsAsync<InvalidDataException>(() => client.CheckEncryptionKeyAsync());
        harness.InvalidPublicKeyId = false; using KsefCredentials credentials = await Authenticate(client, harness);
        using var other = new KsefClient(handler: new KsefHarness(), clock: new KsefClock());
        await Assert.ThrowsAsync<InvalidOperationException>(() => other.OpenOnlineSessionAsync(credentials));
        using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials); KsefSubmission submission = await client.SubmitAsync(credentials, session, InvoiceBytes(), Validator());
        harness.SpoofHash = true; await Assert.ThrowsAsync<InvalidDataException>(() => client.GetInvoiceStatusAsync(credentials, submission));
        clock.Now = KsefHarness.Now.AddHours(2); await Assert.ThrowsAsync<InvalidOperationException>(() => client.GetInvoiceStatusAsync(credentials, submission));
    }
}
