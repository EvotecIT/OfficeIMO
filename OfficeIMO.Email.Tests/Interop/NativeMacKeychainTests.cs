#if NET8_0_OR_GREATER
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Email.Tests;

[CollectionDefinition("Native macOS Keychain", DisableParallelization = true)]
public sealed class NativeMacKeychainCollection { }

[Collection("Native macOS Keychain")]
public sealed class NativeMacKeychainTests(Xunit.Abstractions.ITestOutputHelper output) {
    [NativeMacKeychainFact]
    public void CallerSelectedNonExtractableIdentitySignsDecryptsAndHonorsKeychainAuthorization() {
        if (!OperatingSystem.IsMacOS()) throw new PlatformNotSupportedException();
        byte interaction = MacKeychainFixture.GetInteraction();
        try {
            // CI has no person to authorize access. An unexpected native prompt
            // must become an observable failure rather than a blocked test host.
            MacKeychainFixture.SetInteraction(0);
            QualifyIdentity();
        } finally { MacKeychainFixture.SetInteraction(interaction); }
    }

    private void QualifyIdentity() {
        output.WriteLine("Creating an isolated non-extractable identity.");
        using var fixture = new MacKeychainFixture(Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_KEYCHAIN_SCRATCH")!);
        byte[] content = Encoding.UTF8.GetBytes("OfficeIMO native Keychain — Zażółć 日本語");
        using (X509Certificate2 native = fixture.Open()) {
            using RSA key = native.GetRSAPrivateKey()!;
            Assert.ThrowsAny<CryptographicException>(() => key.ExportParameters(true));
            foreach (string algorithm in new[] { "2.16.840.1.101.3.4.1.2", "2.16.840.1.101.3.4.1.22", "2.16.840.1.101.3.4.1.42" }) {
                foreach (bool ski in new[] { false, true }) {
                    var envelope = new EnvelopedCms(new ContentInfo(content), new AlgorithmIdentifier(new Oid(algorithm)));
                    envelope.Encrypt(new CmsRecipient(ski ? SubjectIdentifierType.SubjectKeyIdentifier : SubjectIdentifierType.IssuerAndSerialNumber,
                        fixture.PublicCertificate));
                    CmsDecryptionResult result = CmsEnvelopedDataService.Decrypt(envelope.Encode(), native);
                    Assert.True(result.Decrypted, string.Join(",", result.Findings.Select(finding => finding.Code)));
                    Assert.Equal(content, result.Content);
                }
            }
            output.WriteLine("Native AES decryption passed; checking signing.");
            var signed = new SignedCms();
            signed.Decode(CmsSignedDataSigner.SignEncapsulated(content, native));
            signed.CheckSignature(verifySignatureOnly: true);
            Assert.Equal(content, signed.ContentInfo.Content);

            var document = new EmailDocument { Subject = "Synthetic Keychain", Body = { Text = "Zażółć 日本語" } };
            EmailSmimeCreationResult message = EmailSmime.SignAndEncrypt(document, native,
                new[] { fixture.PublicCertificate }, OfficeSecurityProvider.Default);
            using EmailReadResult parsed = new EmailDocumentReader().Read(message.Message);
            EmailSmimeProcessingResult processed = EmailSmime.DecryptThenVerify(parsed.Document, native, OfficeSecurityProvider.Default);
            Assert.Equal(document.Body.Text, processed.Content!.Body.Text);
            Assert.True(processed.Verification!.IsCryptographicallyValid);
        }
        using (X509Certificate2 reopened = fixture.Open()) {
            Assert.Equal(fixture.PublicCertificate.RawData, reopened.RawData);
            Assert.True(reopened.HasPrivateKey);
        }
        output.WriteLine("Checking locked-keychain authorization.");
        try {
            fixture.Lock();
            Exception? denied = Record.Exception(() => {
                using X509Certificate2 locked = fixture.Open();
                CmsSignedDataSigner.SignEncapsulated(content, locked);
            });
            Assert.True(denied is CryptographicException or NativeKeychainException,
                "A locked isolated keychain must reject signing without an unlock prompt.");
        } finally { fixture.Unlock(); }
        using (X509Certificate2 authorized = fixture.Open())
            Assert.NotEmpty(CmsSignedDataSigner.SignEncapsulated(content, authorized));
        fixture.Dispose();
        Assert.False(File.Exists(fixture.KeychainPath));
        // Security may retain native identities after backing-file deletion. The
        // caller's path-based selector refuses revoked sources and drops old handles.
        Assert.Throws<FileNotFoundException>(() => fixture.Open());
    }
}

public sealed class NativeMacKeychainFactAttribute : FactAttribute {
    public NativeMacKeychainFactAttribute() {
        if (!OperatingSystem.IsMacOS()) Skip = "Native Keychain qualification requires macOS and .NET 8 or later.";
        else if (Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_NATIVE_KEYCHAIN") != "1" ||
            Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_KEYCHAIN_SCRATCH") is not { } root ||
            !Path.IsPathFullyQualified(root) || !Directory.Exists(root))
            Skip = "Set OFFICEIMO_EMAIL_NATIVE_KEYCHAIN=1 and OFFICEIMO_EMAIL_KEYCHAIN_SCRATCH to an existing task scratch root.";
    }
}

internal sealed class MacKeychainFixture : IDisposable {
    private readonly string _root;
    private readonly string _password = Convert.ToHexString(RandomNumberGenerator.GetBytes(24));
    private readonly string _searchList;
    private bool _created;
    private bool _disposed;
    internal string KeychainPath { get; }
    internal X509Certificate2 PublicCertificate { get; private set; } = null!;

    internal MacKeychainFixture(string scratchRoot) {
        if (!OperatingSystem.IsMacOS()) throw new PlatformNotSupportedException();
        _root = Path.Combine(Path.GetFullPath(scratchRoot), "officeimo-keychain-" + Guid.NewGuid().ToString("N"));
        KeychainPath = Path.Combine(_root, "synthetic.keychain-db");
        _searchList = Run("list-keychains", "-d", "user");
        Directory.CreateDirectory(_root);
        string privateKey = Path.Combine(_root, "synthetic-private.rsa");
        string certificate = Path.Combine(_root, "synthetic.cer");
        try {
            // Drop the generating key before native lookup; there must be no second
            // exportable identity with the same public key in .NET's temporary store.
            using (RSA key = RSA.Create(2048)) {
                var request = new CertificateRequest("CN=OfficeIMO synthetic Keychain qualification", key,
                    HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, false));
                request.CertificateExtensions.Add(new X509KeyUsageExtension(
                    X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment, true));
                using X509Certificate2 generated = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
                byte[] exported = key.ExportRSAPrivateKey();
                try {
                    File.WriteAllBytes(privateKey, exported);
                    File.SetUnixFileMode(privateKey, UnixFileMode.UserRead | UnixFileMode.UserWrite);
                } finally { CryptographicOperations.ZeroMemory(exported); }
                File.WriteAllBytes(certificate, generated.RawData);
#pragma warning disable SYSLIB0057 // .NET 8 has no X509CertificateLoader; this contains only public certificate bytes.
                PublicCertificate = new X509Certificate2(generated.RawData);
#pragma warning restore SYSLIB0057
            }
            Run("create-keychain", "-p", _password, KeychainPath);
            _created = true;
            Unlock();
            // security's PKCS12 route ignores -x. Separate key import uses native
            // key attributes; the test must also prove private export is rejected.
            Run("import", privateKey, "-k", KeychainPath, "-f", "openssl", "-t", "priv", "-P", _password,
                "-x", "-T", Environment.ProcessPath!);
            Run("import", certificate, "-k", KeychainPath, "-f", "x509", "-t", "cert");
        } catch { Dispose(); throw; }
        finally {
            if (File.Exists(privateKey)) File.Delete(privateKey);
            if (File.Exists(certificate)) File.Delete(certificate);
        }
    }

    internal X509Certificate2 Open() {
        if (!File.Exists(KeychainPath)) throw new FileNotFoundException("The caller-selected keychain is unavailable.");
        IntPtr keychain = IntPtr.Zero, search = IntPtr.Zero, identity = IntPtr.Zero;
        try {
            Check(SecKeychainOpen(KeychainPath, out keychain));
            Check(SecIdentitySearchCreate(keychain, 0, out search));
            Check(SecIdentitySearchCopyNext(search, out identity));
#pragma warning disable SYSLIB0057 // Apple's .NET PAL retains an explicitly supplied native SecIdentity handle.
            return new X509Certificate2(identity);
#pragma warning restore SYSLIB0057
        } finally {
            if (identity != IntPtr.Zero) CFRelease(identity);
            if (search != IntPtr.Zero) CFRelease(search);
            if (keychain != IntPtr.Zero) CFRelease(keychain);
        }
    }

    internal void Lock() => Run("lock-keychain", KeychainPath);
    internal void Unlock() => Run("unlock-keychain", "-p", _password, KeychainPath);
    public void Dispose() {
        if (_disposed) return;
        PublicCertificate?.Dispose();
        if (_created) { Run("delete-keychain", KeychainPath); _created = false; }
        // Private-key files can remain if construction failed before import.
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
        _disposed = true;
        if (Run("list-keychains", "-d", "user") != _searchList)
            throw new InvalidOperationException("The user keychain search list changed during isolated qualification.");
    }

    private static string Run(params string[] arguments) {
        var start = new ProcessStartInfo("/usr/bin/security") { RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start)!;
        Task<string> output = process.StandardOutput.ReadToEndAsync(), error = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(15000)) { process.Kill(); throw new IOException("Native security command timed out: " + arguments[0]); }
        Task.WaitAll(output, error);
        if (process.ExitCode != 0) throw new IOException("Native security command failed: " + arguments[0] + " exit " + process.ExitCode);
        return output.Result;
    }

    private const string Security = "/System/Library/Frameworks/Security.framework/Security";
    private const string CoreFoundation = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
    [DllImport(Security)] private static extern int SecKeychainOpen([MarshalAs(UnmanagedType.LPUTF8Str)] string path, out IntPtr keychain);
    [DllImport(Security)] private static extern int SecIdentitySearchCreate(IntPtr keychain, uint keyUsage, out IntPtr search);
    [DllImport(Security)] private static extern int SecIdentitySearchCopyNext(IntPtr search, out IntPtr identity);
    [DllImport(Security)] private static extern int SecKeychainGetUserInteractionAllowed(out byte allowed);
    [DllImport(Security)] private static extern int SecKeychainSetUserInteractionAllowed(byte allowed);
    [DllImport(CoreFoundation)] private static extern void CFRelease(IntPtr value);
    internal static byte GetInteraction() { Check(SecKeychainGetUserInteractionAllowed(out byte result)); return result; }
    internal static void SetInteraction(byte value) => Check(SecKeychainSetUserInteractionAllowed(value));
    private static void Check(int status) { if (status != 0) throw new NativeKeychainException(status); }
}

internal sealed class NativeKeychainException(int status) : Exception("Native synthetic identity status " + status);
#endif
