using OfficeIMO.Invoicing.KSeF;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.KSeF.Tests;

public class KsefLiveTests {
    [PublicEndpointFact]
    public async Task PublicTestEncryptionCertificateContract() {
        using var client = new KsefClient();
        Assert.Equal(44, (await client.CheckEncryptionKeyAsync()).Length);
    }
    [AuthenticatedTestFact]
    public async Task AuthenticatedTestSubmissionStatusAndReceiptQualification() {
        string Setting(string name) => Environment.GetEnvironmentVariable(name) ?? throw new InvalidOperationException("Required live TEST configuration is missing: " + name);
        byte[] ReadBounded(string path, int maximum) {
            using var stream = File.OpenRead(path);
            if (stream.Length <= 0 || stream.Length > maximum) throw new InvalidDataException("Configured TEST input exceeds its supported bound.");
            byte[] bytes = new byte[stream.Length]; stream.ReadExactly(bytes);
            if (stream.ReadByte() != -1) throw new InvalidDataException("Configured TEST input changed while reading.");
            return bytes;
        }
        string tokenPath = Setting("OFFICEIMO_KSEF_TEST_TOKEN_FILE"), invoicePath = Setting("OFFICEIMO_KSEF_TEST_INVOICE_FILE");
        string receiptDirectory = Path.GetFullPath(Setting("OFFICEIMO_KSEF_TEST_RECEIPT_DIRECTORY")); Directory.CreateDirectory(receiptDirectory);
        byte[] tokenBytes = ReadBounded(tokenPath, 32_768);
        using var secret = new KsefSecret(System.Text.Encoding.UTF8.GetString(tokenBytes).TrimEnd('\r', '\n'));
        System.Security.Cryptography.CryptographicOperations.ZeroMemory(tokenBytes);
        var validator = new Fa3SchemaValidator(Fa3SchemaBundle.LoadDirectory(Setting("OFFICEIMO_KSEF_TEST_SCHEMA_DIRECTORY")));
        byte[] xml = ReadBounded(invoicePath, 16 * 1024 * 1024);
        Assert.True(validator.Validate(xml).IsValid, "Configured invoice must pass pinned FA(3) validation before live authentication.");
        using var client = new KsefClient(KsefEnvironment.Test); using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(10));
        using KsefAuthentication authentication = await client.BeginTokenAuthenticationAsync(new KsefContext(KsefContextKind.Nip, Setting("OFFICEIMO_KSEF_TEST_NIP")), secret, timeout.Token);
        async Task<KsefStatus> Wait(Func<Task<KsefStatus>> read) {
            for (int attempt = 0; attempt < 120; attempt++) {
                KsefStatus status = await read(); if (!status.IsPending) return status;
                await Task.Delay(TimeSpan.FromSeconds(2), timeout.Token);
            }
            throw new TimeoutException("KSeF TEST processing remains pending; no mutation is replayed.");
        }
        Assert.True((await Wait(() => client.GetAuthenticationStatusAsync(authentication, timeout.Token))).IsSuccessful);
        using KsefCredentials credentials = await client.RedeemAsync(authentication, timeout.Token);
        using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials, timeout.Token);
        using (var output = new FileStream(Path.Combine(receiptDirectory, session.ReferenceNumber + ".session.json"), FileMode.CreateNew, FileAccess.Write)) {
            System.Text.Json.JsonSerializer.Serialize(output, new { environment = "TEST", context = credentials.Context.Value, session.ReferenceNumber, invoiceHash = Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(xml)) });
        }
        KsefSubmission submission = await client.SubmitAsync(credentials, session, xml, validator, timeout.Token);
        string manifestPath = Path.Combine(receiptDirectory, submission.ReferenceNumber + ".json");
        using (var output = new FileStream(manifestPath, FileMode.CreateNew, FileAccess.Write)) {
            System.Text.Json.JsonSerializer.Serialize(output, new { environment = "TEST", context = credentials.Context.Value, submission.SessionReference, submission.ReferenceNumber, submission.InvoiceHash });
        }
        await client.CloseOnlineSessionAsync(credentials, session, timeout.Token);
        Assert.True((await Wait(async () => (await client.GetInvoiceStatusAsync(credentials, submission, timeout.Token)).Status)).IsSuccessful);
        KsefReceipt receipt = await client.GetInvoiceReceiptAsync(credentials, submission, timeout.Token);
        using (var output = new FileStream(Path.Combine(receiptDirectory, submission.ReferenceNumber + ".upo.xml"), FileMode.CreateNew, FileAccess.Write)) output.Write(receipt.GetBytes());
        Assert.True(receipt.RetrievedThroughAuthenticatedApi); Assert.True(receipt.IsBound);
        Assert.Equal(InvoiceValidationStatus.Passed, receipt.SchemaStatus);
    }
    private sealed class PublicEndpointFactAttribute : FactAttribute {
        public PublicEndpointFactAttribute() { if (Environment.GetEnvironmentVariable("OFFICEIMO_KSEF_PUBLIC_TEST") != "1") Skip = "Opt-in read-only official TEST endpoint proof."; }
    }
    private sealed class AuthenticatedTestFactAttribute : FactAttribute {
        public AuthenticatedTestFactAttribute() { if (Environment.GetEnvironmentVariable("OFFICEIMO_KSEF_LIVE_TEST") != "1") Skip = "No configured TEST identity; opt-in authenticated qualification harness."; }
    }
}
