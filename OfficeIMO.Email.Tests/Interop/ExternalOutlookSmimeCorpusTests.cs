using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Email;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class ExternalOutlookSmimeCorpusTests {
    private const string CorpusVariable = "OFFICEIMO_EMAIL_SMIME_CORPUS";
    private const string CorpusRevision = "HiraokaHyperTools/smime_mail_samples@9e2b7f45e00c98d15ed901b9793fb3c08c20400a";
    private static readonly IReadOnlyDictionary<string, string> ExpectedSha256 =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["certs/smime-2-secret.pfx"] = "381F0A8C49D05CBE45E8D6FCEF9EBF4B101F5156EEDEEA61A2F8474D3CFF2A8E",
            ["received-at-smime2/Message from smime1 to smime2 (signed).eml"] = "4C20FC198493CE7B263B4FD3E5343455B0A2D724C19E1FC44E618D2B19831ECA",
            ["received-at-smime2/Message from smime1 to smime2 (signed).msg"] = "BCF898A0CA49726ED4ABAD909B7566F085B3B9AC1A8830DD5F82B9FBB5739D4F",
            ["received-at-smime2/Message from smime1 to smime2 (encrypted).eml"] = "C1E202A01D6DF10358EACE3F12C8A0B25B220BE911E0FD7003A1D375B071F6C1",
            ["received-at-smime2/Message from smime1 to smime2 (encrypted).msg"] = "32B8393CE2A7D1901B352E03886EA0FA3B6D6A60207C21C9BEF6663705EAC77C",
            ["received-at-smime2/Message from smime1 to smime2 (signed and encrypted).eml"] = "83F8163C8D1579C2408134BD43FE764C03D426BEA2BF607912977D5A351E6182",
            ["received-at-smime2/Message from smime1 to smime2 (signed and encrypted).msg"] = "0F1BD2C7DD869F886E644080CDD156097F5E5D0CC928264CBB193D29990C8C79"
        };

    [Fact]
    public void VerifiesAndDecryptsRealOutlookEmlAndMsgCorpusWhenAvailable() {
        string? root = Environment.GetEnvironmentVariable(CorpusVariable);
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) return;
        AssertExactCorpus(root!);

        string received = Path.Combine(root!, "received-at-smime2");
        string certificatePath = Path.Combine(root!, "certs", "smime-2-secret.pfx");
        Assert.True(Directory.Exists(received), CorpusVariable + " does not identify the expected corpus root.");
        Assert.True(File.Exists(certificatePath), "The corpus recipient certificate is missing.");
        using X509Certificate2 recipient = LoadExportableCertificate(certificatePath, "smime2");

        foreach (string extension in new[] { ".eml", ".msg" }) {
            using EmailReadResult signed = Read(received, "Message from smime1 to smime2 (signed)" + extension);
            EmailSmimeVerificationResult signedResult = EmailSmime.Verify(signed.Document);
            Assert.True(signedResult.IsCryptographicallyValid, Describe(signedResult.Cryptography));
            Assert.Single(signedResult.Cryptography!.Signers);

            using EmailReadResult encrypted = Read(received, "Message from smime1 to smime2 (encrypted)" + extension);
            EmailSmimeDecryptionResult encryptedResult = EmailSmime.Decrypt(encrypted.Document, recipient);
            Assert.True(encryptedResult.Decrypted, Describe(encryptedResult.Cryptography));
            Assert.NotNull(encryptedResult.DecryptedContent);
            Assert.False(encryptedResult.DecryptedContent!.Protection.IsProtected);

            using EmailReadResult nested = Read(
                received,
                "Message from smime1 to smime2 (signed and encrypted)" + extension);
            EmailSmimeDecryptionResult nestedDecryption = EmailSmime.Decrypt(nested.Document, recipient);
            Assert.True(nestedDecryption.Decrypted, Describe(nestedDecryption.Cryptography));
            Assert.NotNull(nestedDecryption.DecryptedContent);
            EmailSmimeVerificationResult nestedSignature = EmailSmime.Verify(nestedDecryption.DecryptedContent!);
            Assert.True(nestedSignature.IsCryptographicallyValid, Describe(nestedSignature.Cryptography));
            Assert.Single(nestedSignature.Cryptography!.Signers);
        }
    }

    private static EmailReadResult Read(string directory, string fileName) {
        string path = Path.Combine(directory, fileName);
        Assert.True(File.Exists(path), "Expected Outlook corpus artifact is missing: " + fileName);
        EmailReadResult read = new EmailDocumentReader().Read(path);
        try {
            Assert.DoesNotContain(read.Diagnostics, diagnostic =>
                diagnostic.Severity == EmailDiagnosticSeverity.Error);
            if (path.EndsWith(".msg", StringComparison.OrdinalIgnoreCase)) {
                Assert.Equal(EmailFileFormat.OutlookMsg, read.Document.Format);
                Assert.StartsWith("IPM.Note.SMIME", read.Document.MessageClass, StringComparison.OrdinalIgnoreCase);
            } else {
                Assert.Equal(EmailFileFormat.Eml, read.Document.Format);
                Assert.Contains(read.Document.Headers, header =>
                    header.Name.Equals("X-Mailer", StringComparison.OrdinalIgnoreCase) &&
                    header.Value.Contains("Microsoft Outlook", StringComparison.OrdinalIgnoreCase));
            }
            return read;
        } catch {
            read.Dispose();
            throw;
        }
    }

    private static X509Certificate2 LoadExportableCertificate(string path, string password) {
#if NET9_0_OR_GREATER
        return X509CertificateLoader.LoadPkcs12FromFile(
            path,
            password,
            X509KeyStorageFlags.Exportable);
#else
        return new X509Certificate2(path, password, X509KeyStorageFlags.Exportable);
#endif
    }

    private static void AssertExactCorpus(string root) {
        foreach (KeyValuePair<string, string> expected in ExpectedSha256) {
            string path = Path.Combine(root, expected.Key.Replace('/', Path.DirectorySeparatorChar));
            Assert.True(File.Exists(path), CorpusRevision + " artifact is missing: " + expected.Key);
            using FileStream stream = File.OpenRead(path);
            using SHA256 sha256 = SHA256.Create();
            string actual = BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", string.Empty);
            Assert.Equal(expected.Value, actual, ignoreCase: true);
        }
    }

    private static string Describe(CmsVerificationResult? result) => result == null
        ? "No CMS verification result was returned."
        : string.Join(" | ", result.Findings.Select(finding => finding.Code + ": " + finding.Message)
            .Concat(result.Signers.SelectMany(signer => signer.Findings)
                .Select(finding => finding.Code + ": " + finding.Message)));

    private static string Describe(CmsDecryptionResult? result) => result == null
        ? "No CMS decryption result was returned."
        : string.Join(" | ", result.Findings.Select(finding => finding.Code + ": " + finding.Message));
}
