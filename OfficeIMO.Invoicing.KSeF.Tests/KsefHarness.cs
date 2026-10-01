using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Invoicing.KSeF;

namespace OfficeIMO.Invoicing.KSeF.Tests;

internal sealed class KsefHarness : HttpMessageHandler {
    internal const string AuthenticationReference = "20260930-AU-0000000001-0000000002-01";
    internal const string SessionReference = "20260930-SO-0000000001-0000000002-01";
    internal const string InvoiceReference = "20260930-II-0000000001-0000000002-01";
    internal const string KsefNumber = "9999999999-20260930-ABCDEF012345-AA";
    internal static readonly DateTimeOffset Now = new(2026, 9, 30, 12, 0, 0, TimeSpan.Zero);
    internal readonly RSA Rsa = RSA.Create(2048);
    private readonly byte[] _certificate;
    internal readonly List<string> Requests = new();
    internal byte[]? Plaintext, SessionKey, Iv;
    internal string? PlaintextHash;
    internal string? ExpectedToken;
    internal Action? OnSubmission;
    internal string? Failure;
    internal bool SpoofHash;
    internal bool InvalidPublicKeyId;
    internal string? RedemptionFailure, CloseFailure;
    internal bool RedemptionPending;
    internal int InvoiceReads;
    internal bool PendingInvoice;
    internal TaskCompletionSource? SubmissionStarted, ReleaseSubmission;

    internal KsefHarness() {
        var request = new CertificateRequest("CN=OfficeIMO credential-free harness", Rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using X509Certificate2 certificate = request.CreateSelfSigned(Now.AddDays(-1), Now.AddYears(1));
        _certificate = certificate.Export(X509ContentType.Cert);
    }
    internal HttpResponseMessage Json(object value, HttpStatusCode code = HttpStatusCode.OK) => new(code) { Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json") };
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        Assert.Equal("api-test.ksef.mf.gov.pl", request.RequestUri!.Host);
        string path = request.RequestUri.AbsolutePath.Substring(4);
        Requests.Add(request.Method + " " + path);
        using JsonDocument? document = request.Content == null ? null : JsonDocument.Parse(await request.Content.ReadAsByteArrayAsync(cancellationToken));
        JsonElement root = document?.RootElement ?? default;
        switch (path) {
            case "security/public-key-certificates":
                Assert.Null(request.Headers.Authorization);
                return Json(new[] { new { certificate = Convert.ToBase64String(_certificate), certificateId = Convert.ToBase64String(SHA256.HashData(_certificate)), publicKeyId = InvalidPublicKeyId ? Convert.ToBase64String(new byte[32]) : Convert.ToBase64String(SHA256.HashData(Rsa.ExportSubjectPublicKeyInfo())), validFrom = Now.AddDays(-1), validTo = Now.AddYears(1), usage = new[] { "KsefTokenEncryption", "SymmetricKeyEncryption" } } });
            case "auth/challenge":
                Assert.Equal(HttpMethod.Post, request.Method); Assert.Null(request.Headers.Authorization);
                return Json(new { challenge = "20260930-CR-0000000001-0000000002-01", timestamp = Now, timestampMs = Now.ToUnixTimeMilliseconds(), clientIp = "127.0.0.1" });
            case "auth/ksef-token":
                Assert.Null(request.Headers.Authorization);
                byte[] token = Rsa.Decrypt(Convert.FromBase64String(root.GetProperty("encryptedToken").GetString()!), RSAEncryptionPadding.OaepSHA256);
                Assert.Equal(ExpectedToken + "|" + Now.ToUnixTimeMilliseconds(), Encoding.UTF8.GetString(token)); CryptographicOperations.ZeroMemory(token);
                Assert.Equal("Nip", root.GetProperty("contextIdentifier").GetProperty("type").GetString());
                return Json(new { referenceNumber = AuthenticationReference, authenticationToken = new { token = "temporary-token", validUntil = Now.AddMinutes(10) } });
            case "auth/" + AuthenticationReference:
                Assert.Equal("temporary-token", request.Headers.Authorization?.Parameter);
                return Json(new { status = new { code = RedemptionPending ? 100 : 200, description = "Authentication" } });
            case "auth/token/redeem":
                Assert.Equal("temporary-token", request.Headers.Authorization?.Parameter);
                if (RedemptionFailure != null) return FailedResponse(RedemptionFailure, request, 1024 * 1024);
                return Json(new { accessToken = new { token = "access-token", validUntil = Now.AddHours(1) }, refreshToken = new { token = "refresh-token", validUntil = Now.AddDays(1) } });
            case "auth/token/refresh":
                Assert.Equal("refresh-token", request.Headers.Authorization?.Parameter);
                return Json(new { accessToken = new { token = "refreshed-access", validUntil = Now.AddHours(2) } });
            case "sessions/online":
                Assert.Contains(request.Headers.Authorization?.Parameter, new[] { "access-token", "refreshed-access" });
                Assert.Equal("FA (3)", root.GetProperty("formCode").GetProperty("systemCode").GetString());
                JsonElement encryption = root.GetProperty("encryption");
                SessionKey = Rsa.Decrypt(Convert.FromBase64String(encryption.GetProperty("encryptedSymmetricKey").GetString()!), RSAEncryptionPadding.OaepSHA256);
                Iv = Convert.FromBase64String(encryption.GetProperty("initializationVector").GetString()!); Assert.Equal(32, SessionKey.Length); Assert.Equal(16, Iv.Length);
                return Json(new { referenceNumber = SessionReference, validUntil = Now.AddHours(12) });
            case "sessions/online/" + SessionReference + "/invoices":
                SubmissionStarted?.TrySetResult();
                if (ReleaseSubmission != null) await ReleaseSubmission.Task.WaitAsync(cancellationToken);
                OnSubmission?.Invoke();
                byte[] encrypted = Convert.FromBase64String(root.GetProperty("encryptedInvoiceContent").GetString()!);
                using (Aes aes = Aes.Create()) { aes.Key = SessionKey!; Plaintext = aes.DecryptCbc(encrypted, Iv!, PaddingMode.PKCS7); }
                PlaintextHash = Convert.ToBase64String(SHA256.HashData(Plaintext));
                Assert.Equal(PlaintextHash, root.GetProperty("invoiceHash").GetString()); Assert.Equal(Plaintext.Length, root.GetProperty("invoiceSize").GetInt32());
                Assert.Equal(Convert.ToBase64String(SHA256.HashData(encrypted)), root.GetProperty("encryptedInvoiceHash").GetString()); Assert.Equal(encrypted.Length, root.GetProperty("encryptedInvoiceSize").GetInt32());
                Assert.False(root.GetProperty("offlineMode").GetBoolean());
                if (Failure == "cancel") throw new OperationCanceledException();
                if (Failure == "network") throw new HttpRequestException("Synthetic connection loss");
                if (Failure == "server") return new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent("raw-sensitive-server-body") };
                if (Failure is "malformed" or "declared-size" or "stream-size" or "redirected-response") return FailedResponse(Failure, request, 1024 * 1024);
                if (Failure == "reject") return new HttpResponseMessage(HttpStatusCode.BadRequest) { Content = new StringContent("raw-sensitive-server-body") };
                return Json(new { referenceNumber = InvoiceReference });
            case "sessions/" + SessionReference + "/invoices/" + InvoiceReference:
                InvoiceReads++;
                return Json(InvoiceStatus());
            case "sessions/" + SessionReference + "/invoices":
                return Json(new { invoices = new[] { InvoiceStatus() }, continuationToken = (string?)null });
            case "sessions/" + SessionReference + "/invoices/" + InvoiceReference + "/upo":
                Assert.Contains(request.Headers.Authorization?.Parameter, new[] { "access-token", "refreshed-access" });
                XDocument receipt = XDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "upo-invoice-nip.xml"));
                void Set(string name, string value) => receipt.Descendants().Single(element => element.Name.LocalName == name).Value = value;
                Set("NazwaPodmiotuPrzyjmujacego", "Ministerstwo Finansów"); Set("NumerReferencyjnySesji", SessionReference); Set("Nip", "9999999999");
                Set("NumerKSeFDokumentu", KsefNumber); Set("NipSprzedawcy", "9999999999"); Set("SkrotDokumentu", PlaintextHash!);
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(receipt.ToString(), Encoding.UTF8, "application/xml") };
            case "sessions/online/" + SessionReference + "/close": return CloseFailure == null ? new HttpResponseMessage(HttpStatusCode.NoContent) : FailedResponse(CloseFailure, request, 65_536);
            case "sessions/" + SessionReference: return Json(new { status = new { code = 200, description = "Closed" }, dateCreated = Now, dateUpdated = Now });
        }
        throw new InvalidOperationException("Unexpected synthetic KSeF route: " + path);
    }
    private static HttpResponseMessage FailedResponse(string failure, HttpRequestMessage request, int maximum) {
        var response = new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(failure == "malformed" ? "{" : "{}") };
        if (failure == "declared-size") response.Content.Headers.ContentLength = maximum + 1;
        if (failure == "stream-size") { response.Content.Dispose(); response.Content = new UnknownLengthContent(maximum + 1); }
        if (failure == "redirected-response") response.RequestMessage = new HttpRequestMessage(request.Method, "https://untrusted.invalid/");
        return response;
    }
    private sealed class UnknownLengthContent(int length) : HttpContent {
        protected override bool TryComputeLength(out long result) { result = 0; return false; }
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) => stream.WriteAsync(new byte[length]).AsTask();
        protected override Task<Stream> CreateContentReadStreamAsync() => Task.FromResult<Stream>(new MemoryStream(new byte[length], writable: false));
    }
    private object InvoiceStatus() => new { ordinalNumber = 1, referenceNumber = InvoiceReference, invoiceHash = SpoofHash ? Convert.ToBase64String(new byte[32]) : PlaintextHash!, invoicingDate = Now, ksefNumber = PendingInvoice ? null : KsefNumber, status = new { code = PendingInvoice ? 100 : 200, description = "Invoice" }, upoDownloadUrl = "http://127.0.0.1:1/untrusted" };
    protected override void Dispose(bool disposing) { if (disposing) { Rsa.Dispose(); if (SessionKey != null) CryptographicOperations.ZeroMemory(SessionKey); } base.Dispose(disposing); }
}
internal sealed class KsefClock : TimeProvider {
    internal DateTimeOffset Now = KsefHarness.Now;
    public override DateTimeOffset GetUtcNow() => Now;
}
