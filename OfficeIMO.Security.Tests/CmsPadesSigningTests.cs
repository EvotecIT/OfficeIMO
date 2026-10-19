using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.Ess;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Tsp;
using Org.BouncyCastle.Utilities.Collections;

namespace OfficeIMO.Security.Tests;

public sealed class CmsPadesSigningTests {
    [Theory]
    [InlineData(false, "SHA256")]
    [InlineData(true, "SHA256")]
    [InlineData(true, "SHA384")]
    [InlineData(true, "SHA512")]
    public void PadesBindsSignerCertificateAndOmitsCmsSigningTime(bool ellipticCurve, string digest) {
        using X509Certificate2 certificate = CreateCertificate(ellipticCurve);
        byte[] data = Encoding.UTF8.GetBytes("PAdES baseline signature");
        byte[] encoded = CmsSignedDataSigner.SignDetached(data, certificate, new CmsSigningOptions {
            UsePadesBaseline = true, DigestAlgorithm = new HashAlgorithmName(digest)
        });
        var cms = new CmsSignedData(new CmsProcessableByteArray(data), encoded);
        SignerInformation signer = Assert.Single(cms.GetSignerInfos().GetSigners());
        Assert.True(signer.Verify(DotNetUtilities.FromX509Certificate(certificate)));
        Assert.Null(signer.SignedAttributes[CmsAttributes.SigningTime]);
        var binding = SigningCertificateV2.GetInstance(signer.SignedAttributes[
            new DerObjectIdentifier("1.2.840.113549.1.9.16.2.47")].AttrValues[0]);
        Assert.Equal(SHA256.HashData(certificate.RawData), Assert.Single(binding.GetCerts()).GetCertHash());
        var platform = new System.Security.Cryptography.Pkcs.SignedCms(
            new System.Security.Cryptography.Pkcs.ContentInfo(data), detached: true);
        platform.Decode(encoded);
        platform.CheckSignature(verifySignatureOnly: true);
    }

    [Fact]
    public void PadesRejectsEncapsulatedContentAndLegacyDigest() {
        using X509Certificate2 certificate = CreateCertificate(false);
        Assert.Throws<ArgumentException>(() => CmsSignedDataSigner.SignEncapsulated(new byte[1], certificate,
            new CmsSigningOptions { UsePadesBaseline = true }));
        Assert.Throws<ArgumentException>(() => CmsSignedDataSigner.SignDetached(new byte[1], certificate,
            new CmsSigningOptions { UsePadesBaseline = true, DigestAlgorithm = HashAlgorithmName.SHA1 }));
    }

    [Fact]
    public async Task TimestampClientAndCmsCallbackProduceBoundVerifiedTimestamp() {
        using X509Certificate2 certificate = CreateCertificate(false, tsa: true);
        var trust = new CertificateValidationOptions { ChainEvaluator = (_, _) => true };
        var client = new Rfc3161TimestampClient((request, _) => Task.FromResult(CreateResponse(request, certificate)), trust);
        byte[] data = Encoding.UTF8.GetBytes("timestamp client");
        byte[] token = await client.GetTimestampAsync(data);
        Assert.Equal(SecurityValidationStatus.Valid, Rfc3161TimestampVerifier.Verify(token, data, trust).Status);
        using X509Certificate2 signerCertificate = CreateCertificate(true);
        byte[] cms = CmsSignedDataSigner.SignDetached(data, signerCertificate, new CmsSigningOptions {
            UsePadesBaseline = true, SignatureTimestampProvider = client.GetTimestamp
        });
        SignerInformation signer = Assert.Single(new CmsSignedData(cms).GetSignerInfos().GetSigners());
        var timestamp = signer.UnsignedAttributes[new DerObjectIdentifier("1.2.840.113549.1.9.16.2.14")];
        byte[] embedded = timestamp.AttrValues[0].GetEncoded();
        Assert.Equal(SecurityValidationStatus.Valid, Rfc3161TimestampVerifier.Verify(embedded, signer.GetSignature(), trust).Status);
        Assert.Throws<CryptographicException>(() => CmsSignedDataSigner.SignDetached(data, signerCertificate,
            new CmsSigningOptions { SignatureTimestampProvider = _ => token }));
    }

    [Fact]
    public async Task TimestampClientRejectsReplayedUntrustedOversizedAndCancelledResponses() {
        using X509Certificate2 certificate = CreateCertificate(false, tsa: true);
        byte[]? previousResponse = null;
        var trust = new CertificateValidationOptions { ChainEvaluator = (_, _) => true };
        var client = new Rfc3161TimestampClient((request, _) => Task.FromResult(previousResponse ??= CreateResponse(request, certificate)), trust);
        await client.GetTimestampAsync(new byte[] { 1 });
        await Assert.ThrowsAsync<TspValidationException>(() => client.GetTimestampAsync(new byte[] { 1 }));
        var untrusted = new Rfc3161TimestampClient((request, _) => Task.FromResult(CreateResponse(request, certificate)),
            new CertificateValidationOptions { ChainEvaluator = (_, _) => false });
        await Assert.ThrowsAsync<CryptographicException>(() => untrusted.GetTimestampAsync(new byte[1]));
        var oversized = new Rfc3161TimestampClient((_, _) => Task.FromResult(new byte[33]), trust, maximumResponseBytes: 32);
        await Assert.ThrowsAsync<ArgumentException>(() => oversized.GetTimestampAsync(new byte[1]));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => client.GetTimestampAsync(new byte[1], cancellation.Token));
    }

    [Fact]
    public async Task SynchronousTimestampBridgeDoesNotBlockAnAsyncTransportOnTheCallerContext() {
        using X509Certificate2 certificate = CreateCertificate(false, tsa: true);
        var trust = new CertificateValidationOptions { ChainEvaluator = (_, _) => true };
        var client = new Rfc3161TimestampClient(async (request, _) => {
            await Task.Yield(); // A normal HTTP await also captures a caller's UI context.
            return CreateResponse(request, certificate);
        }, trust);
        var context = new HeldSynchronizationContext();
        var completion = new TaskCompletionSource<byte[]>(TaskCreationOptions.RunContinuationsAsynchronously);
        byte[] data = Encoding.UTF8.GetBytes("UI timestamp request");
        var caller = new Thread(() => {
            SynchronizationContext.SetSynchronizationContext(context);
            try { completion.SetResult(client.GetTimestamp(data)); }
            catch (Exception error) { completion.SetException(error); }
        }) { IsBackground = true };
        caller.Start();
        try {
            byte[] token = await completion.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(SecurityValidationStatus.Valid, Rfc3161TimestampVerifier.Verify(token, data, trust).Status);
        } finally {
            // Drain captured work even on failure, so the regression cannot strand a thread.
            context.Release();
            Assert.True(caller.Join(TimeSpan.FromSeconds(5)), "Timestamp caller did not finish after releasing its context.");
        }
    }

    private sealed class HeldSynchronizationContext : SynchronizationContext {
        private readonly object _sync = new();
        private readonly List<(SendOrPostCallback Callback, object? State)> _pending = new();
        private bool _released;
        public override void Post(SendOrPostCallback callback, object? state) {
            lock (_sync) {
                if (!_released) { _pending.Add((callback, state)); return; }
            }
            ThreadPool.QueueUserWorkItem(_ => callback(state));
        }
        internal void Release() {
            lock (_sync) {
                _released = true;
                foreach (var pending in _pending)
                    ThreadPool.QueueUserWorkItem(_ => pending.Callback(pending.State));
                _pending.Clear();
            }
        }
    }

    internal static X509Certificate2 CreateCertificate(bool ellipticCurve, bool tsa = false) {
        using RSA? rsa = ellipticCurve ? null : RSA.Create(2048);
        using ECDsa? ec = ellipticCurve ? ECDsa.Create(ECCurve.NamedCurves.nistP256) : null;
        var request = ec is null
            ? new CertificateRequest("CN=OfficeIMO ephemeral signer", rsa!, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1)
            : new CertificateRequest("CN=OfficeIMO ephemeral EC signer", ec, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        if (tsa) request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.8") }, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    internal static byte[] CreateResponse(byte[] requestBytes, X509Certificate2 certificate) {
        using RSA rsa = certificate.GetRSAPrivateKey()!;
        var bcCertificate = DotNetUtilities.FromX509Certificate(certificate);
        var factory = new Org.BouncyCastle.Crypto.Operators.Asn1SignatureFactory("SHA256WITHRSA", DotNetUtilities.GetRsaKeyPair(rsa).Private);
        var signer = new SignerInfoGeneratorBuilder().Build(factory, bcCertificate);
        var tokens = new TimeStampTokenGenerator(signer, Org.BouncyCastle.Crypto.Operators.Asn1DigestFactory.Get("SHA256"),
            new DerObjectIdentifier("1.3.6.1.4.1.59069.1.1"), true);
        tokens.SetCertificates(new CertificateStore(bcCertificate));
        var responses = new TimeStampResponseGenerator(tokens, TspAlgorithms.Allowed);
        return responses.Generate(new TimeStampRequest(requestBytes), Org.BouncyCastle.Math.BigInteger.One, DateTime.UtcNow).GetEncoded();
    }

    private sealed class CertificateStore(Org.BouncyCastle.X509.X509Certificate certificate) : IStore<Org.BouncyCastle.X509.X509Certificate> {
        public IEnumerable<Org.BouncyCastle.X509.X509Certificate> EnumerateMatches(ISelector<Org.BouncyCastle.X509.X509Certificate> selector) {
            if (selector is null || selector.Match(certificate)) yield return certificate;
        }
    }
}
