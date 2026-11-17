using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Operators;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Ocsp;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Tsp;
using Org.BouncyCastle.Utilities.Collections;
using Org.BouncyCastle.X509;
using BcCertificate = Org.BouncyCastle.X509.X509Certificate;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Throwaway test PKI: a root CA, signer, time-stamp authority and delegated OCSP responder, with OCSP and CRL
/// responders served over HTTP on the loopback interface. Clocks are explicit so tests can sign in the past.
/// </summary>
internal sealed class PdfTestPki : IDisposable {
    private static readonly SecureRandom Random = new();
    private readonly AsymmetricCipherKeyPair _rootKey = CreateKey();
    private readonly AsymmetricCipherKeyPair _tsaKey = CreateKey();
    private readonly AsymmetricCipherKeyPair _responderKey = CreateKey();
    private readonly ConcurrentDictionary<string, DateTime> _revoked = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, BcCertificate> _issued = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, AsymmetricKeyParameter> _signerKeys = new(StringComparer.OrdinalIgnoreCase);
    private readonly TcpListener _listener;
    private readonly CancellationTokenSource _stop = new();
    private readonly Thread _serverThread;
    private readonly object _clientGate = new();
    private TcpClient? _activeClient;
    private Exception? _serverError;
    private readonly bool _ocsp;
    private readonly bool _crl;
    private readonly bool _crlCoversOnlyAuthorities;
    private readonly bool _crlUnknownCriticalExtension;
    private readonly BcCertificate _rootCertificate;
    private readonly BcCertificate _tsaCertificate;
    private readonly BcCertificate? _responderCertificate;
    private long _serial = 100;
    private int _ocspRequests;
    private int _crlRequests;

    internal PdfTestPki(DateTime now, bool ocsp = true, bool crl = true, bool delegatedResponder = true, DateTime? responderNotAfter = null,
        bool crlCoversOnlyAuthorities = false, bool crlUnknownCriticalExtension = false) {
        _crlCoversOnlyAuthorities = crlCoversOnlyAuthorities;
        _crlUnknownCriticalExtension = crlUnknownCriticalExtension;
        Now = now;
        ResponderTime = now;
        _ocsp = ocsp;
        _crl = crl;
        _listener = new TcpListener(IPAddress.Loopback, 0);
        _listener.Start();
        BaseUrl = "http://127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port + "/";

        var rootName = new X509Name("CN=OfficeIMO revocation test root");
        _rootCertificate = Issue(rootName, rootName, _rootKey.Public, _rootKey.Private, now.AddDays(-365), now.AddDays(365), generator => {
            generator.AddExtension(X509Extensions.BasicConstraints, true, new BasicConstraints(true));
            generator.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.KeyCertSign | KeyUsage.CrlSign));
        });
        _tsaCertificate = IssueLeaf("CN=OfficeIMO revocation test TSA", _tsaKey.Public, now.AddDays(-60), now.AddDays(60), generator => {
            generator.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.DigitalSignature));
            generator.AddExtension(X509Extensions.ExtendedKeyUsage, true,
                new ExtendedKeyUsage(new List<DerObjectIdentifier> { new("1.3.6.1.5.5.7.3.8") }));
        });
        if (delegatedResponder) {
            _responderCertificate = Issue(new X509Name("CN=OfficeIMO revocation test OCSP responder"), rootName, _responderKey.Public,
                _rootKey.Private, now.AddDays(-60), responderNotAfter ?? now.AddDays(60), generator => {
                    generator.AddExtension(X509Extensions.BasicConstraints, true, new BasicConstraints(false));
                    generator.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.DigitalSignature));
                    generator.AddExtension(X509Extensions.ExtendedKeyUsage, false,
                        new ExtendedKeyUsage(new List<DerObjectIdentifier> { new("1.3.6.1.5.5.7.3.9") }));
                    generator.AddExtension(OcspObjectIdentifiers.PkixOcspNocheck, false, DerNull.Instance);
                });
        }
        Root = Load(_rootCertificate.GetEncoded());
        Tsa = ToPlatform(_tsaCertificate, _tsaKey);
        // The platform can synchronously fetch revocation data while worker threads are occupied.
        _serverThread = new Thread(Serve) { IsBackground = true, Name = "OfficeIMO test PKI responder" };
        _serverThread.Start();
    }

    internal DateTime Now { get; }

    /// <summary>RFC 6960 archive cutoff added to every OCSP single response, when set.</summary>
    internal DateTime? OcspArchiveCutoff { get; set; }

    /// <summary>Time at which OCSP responses and CRLs are produced.</summary>
    internal DateTime ResponderTime { get; set; }

    internal string BaseUrl { get; }
    internal X509Certificate2 Root { get; }
    internal X509Certificate2 Tsa { get; }
    internal int OcspRequests => Volatile.Read(ref _ocspRequests);
    internal int CrlRequests => Volatile.Read(ref _crlRequests);

    internal X509Certificate2 IssueSigner(string name, DateTime notBefore, DateTime notAfter) {
        AsymmetricCipherKeyPair key = CreateKey();
        BcCertificate certificate = IssueLeaf("CN=" + name, key.Public, notBefore, notAfter, generator =>
            generator.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.DigitalSignature | KeyUsage.NonRepudiation)));
        X509Certificate2 platform = ToPlatform(certificate, key);
        _signerKeys[platform.Thumbprint] = key.Private;
        return platform;
    }

    /// <summary>Private key of a signer issued by this PKI, for building custom CMS without exporting a platform key.</summary>
    internal AsymmetricKeyParameter PrivateKeyOf(X509Certificate2 signer) => _signerKeys[signer.Thumbprint];

    internal void Revoke(X509Certificate2 certificate, DateTime revokedAt) =>
        _revoked[new BigInteger(1, Enumerable.Reverse(certificate.GetSerialNumber()).ToArray()).ToString(16)] = revokedAt;

    /// <summary>Trusts only this PKI's root, with no platform download or revocation retrieval.</summary>
    internal CmsVerificationOptions VerificationOptions() {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.DisableCertificateDownloads = false;
        options.CertificateValidation.ExtraCertificates.Add(Root);
        string rootThumbprint = Root.Thumbprint;
        options.CertificateValidation.ChainEvaluator = (_, chain) =>
            chain.ChainElements.Count > 0 &&
            string.Equals(chain.ChainElements[chain.ChainElements.Count - 1].Certificate.Thumbprint, rootThumbprint, StringComparison.OrdinalIgnoreCase);
        return options;
    }

    internal PdfCmsSignatureCryptographyProvider Provider() =>
        new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default, VerificationOptions());

    internal Rfc3161TimestampClient TimestampClient(DateTime generationTime) =>
        new Rfc3161TimestampClient((request, _) => Task.FromResult(CreateTimestampResponse(request, generationTime)),
            VerificationOptions().CertificateValidation);

    internal byte[] Sign(X509Certificate2 signer, DateTime? timestampTime, PdfSignatureProfile profile = PdfSignatureProfile.Approval,
        PdfCertificationPermissionLevel permission = PdfCertificationPermissionLevel.NoChanges) {
        var options = new CmsSigningOptions();
        if (timestampTime.HasValue) options.SignatureTimestampProvider = TimestampClient(timestampTime.Value).GetTimestamp;
        using var cmsSigner = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, signer,
            signingOptions: options, certificateChain: new[] { Root });
        return PdfDocument.Create().Meta(title: "OfficeIMO revocation test")
            .Paragraph(paragraph => paragraph.AddText("Long-term validation"))
            .Security.SignExternal(cmsSigner, new PdfExternalSignatureOptions {
                SubFilter = PdfExternalSignatureSubFilter.CadesDetached,
                Profile = profile,
                CertificationPermission = permission,
                FieldName = "Signature1",
                ReservedSignatureContentsBytes = 24576
            }).ToDocument().ToBytes();
    }

    internal byte[] SignAgain(byte[] pdf, X509Certificate2 signer, DateTime timestampTime) {
        var options = new CmsSigningOptions { SignatureTimestampProvider = TimestampClient(timestampTime).GetTimestamp };
        using var cmsSigner = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, signer,
            signingOptions: options, certificateChain: new[] { Root });
        return PdfDocument.Load(pdf).Security.SignExternal(cmsSigner, new PdfExternalSignatureOptions {
            SubFilter = PdfExternalSignatureSubFilter.CadesDetached,
            FieldName = "Signature2",
            ReservedSignatureContentsBytes = 24576
        }).ToDocument().ToBytes();
    }

    internal byte[] AddDocumentTimestamp(byte[] pdf, DateTime generationTime) =>
        PdfDocument.Load(pdf).Security.SignExternal(new DocumentTimestampSigner(TimestampClient(generationTime)),
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.DocumentTimestamp,
                FieldName = "DocumentTimestamp" + generationTime.Ticks,
                ReservedSignatureContentsBytes = 24576
            }).ToDocument().ToBytes();

    public void Dispose() {
        _stop.Cancel();
        _listener.Stop();
        lock (_clientGate) {
            _activeClient?.Dispose();
        }
        try {
            if (!_serverThread.Join(TimeSpan.FromSeconds(5))) {
                throw new TimeoutException("The test PKI responder did not stop.");
            }
            if (_serverError is not null) {
                throw new InvalidOperationException("The test PKI responder failed.", _serverError);
            }
        } finally {
            _stop.Dispose();
            Root.Dispose();
            Tsa.Dispose();
        }
    }

    private BcCertificate IssueLeaf(string subject, AsymmetricKeyParameter publicKey, DateTime notBefore, DateTime notAfter,
        Action<X509V3CertificateGenerator> extensions) =>
        Issue(new X509Name(subject), _rootCertificate.SubjectDN, publicKey, _rootKey.Private, notBefore, notAfter, generator => {
            generator.AddExtension(X509Extensions.BasicConstraints, true, new BasicConstraints(false));
            if (_ocsp) {
                generator.AddExtension(X509Extensions.AuthorityInfoAccess, false, new AuthorityInformationAccess(
                    new AccessDescription(AccessDescription.IdADOcsp, new GeneralName(GeneralName.UniformResourceIdentifier, BaseUrl + "ocsp"))));
            }
            if (_crl) {
                generator.AddExtension(X509Extensions.CrlDistributionPoints, false, new CrlDistPoint(new[] {
                    new DistributionPoint(new DistributionPointName(new GeneralNames(
                        new GeneralName(GeneralName.UniformResourceIdentifier, BaseUrl + "root.crl"))), null, null)
                }));
            }
            extensions(generator);
        });

    private BcCertificate Issue(X509Name subject, X509Name issuer, AsymmetricKeyParameter publicKey, AsymmetricKeyParameter signingKey,
        DateTime notBefore, DateTime notAfter, Action<X509V3CertificateGenerator> extensions) {
        var generator = new X509V3CertificateGenerator();
        generator.SetSerialNumber(BigInteger.ValueOf(Interlocked.Increment(ref _serial)));
        generator.SetIssuerDN(issuer);
        generator.SetSubjectDN(subject);
        generator.SetNotBefore(notBefore);
        generator.SetNotAfter(notAfter);
        generator.SetPublicKey(publicKey);
        extensions(generator);
        BcCertificate certificate = generator.Generate(new Asn1SignatureFactory("SHA256WITHRSA", signingKey, Random));
        _issued[certificate.SerialNumber.ToString(16)] = certificate;
        return certificate;
    }

    private byte[] CreateOcspResponse(byte[] encodedRequest) {
        var request = new OcspReq(encodedRequest);
        var generator = new BasicOcspRespGenerator(new RespID((_responderCertificate ?? _rootCertificate).SubjectDN));
        DateTime thisUpdate = ResponderTime;
        foreach (Req item in request.GetRequestList()) {
            CertificateID id = item.GetCertID();
            CertificateStatus? status = _revoked.TryGetValue(id.SerialNumber.ToString(16), out DateTime revokedAt)
                ? new RevokedStatus(revokedAt, CrlReason.KeyCompromise)
                : _issued.ContainsKey(id.SerialNumber.ToString(16)) ? CertificateStatus.Good : new UnknownStatus();
            X509Extensions? extensions = OcspArchiveCutoff is DateTime cutoff
                ? new X509Extensions(new Dictionary<DerObjectIdentifier, Org.BouncyCastle.Asn1.X509.X509Extension> {
                    [new DerObjectIdentifier("1.3.6.1.5.5.7.48.1.6")] = new Org.BouncyCastle.Asn1.X509.X509Extension(false, new DerOctetString(new DerGeneralizedTime(cutoff)))
                })
                : null;
            generator.AddResponse(id, status, thisUpdate, thisUpdate.AddDays(7), extensions);
        }
        AsymmetricKeyParameter key = _responderCertificate is null ? _rootKey.Private : _responderKey.Private;
        BcCertificate[] chain = _responderCertificate is null ? Array.Empty<BcCertificate>() : new[] { _responderCertificate };
        BasicOcspResp basic = generator.Generate(new Asn1SignatureFactory("SHA256WITHRSA", key, Random), chain, thisUpdate);
        return new OCSPRespGenerator().Generate(OCSPRespGenerator.Successful, basic).GetEncoded();
    }

    private byte[] CreateCrl() {
        var generator = new X509V2CrlGenerator();
        generator.SetIssuerDN(_rootCertificate.SubjectDN);
        generator.SetThisUpdate(ResponderTime);
        generator.SetNextUpdate(ResponderTime.AddDays(7));
        foreach (KeyValuePair<string, DateTime> revoked in _revoked) {
            generator.AddCrlEntry(new BigInteger(revoked.Key, 16), revoked.Value, CrlReason.KeyCompromise);
        }
        generator.AddExtension(X509Extensions.CrlNumber, false, new CrlNumber(BigInteger.ValueOf(ResponderTime.Ticks)));
        if (_crlUnknownCriticalExtension) {
            generator.AddExtension(new DerObjectIdentifier("1.3.6.1.4.1.59069.9.1"), true, DerNull.Instance);
        }
        if (_crlCoversOnlyAuthorities) {
            generator.AddExtension(X509Extensions.IssuingDistributionPoint, true,
                new IssuingDistributionPoint(null, false, true, null, false, false));
        }
        return generator.Generate(new Asn1SignatureFactory("SHA256WITHRSA", _rootKey.Private, Random)).GetEncoded();
    }

    private byte[] CreateTimestampResponse(byte[] request, DateTime generationTime) {
        SignerInfoGenerator signer = new SignerInfoGeneratorBuilder()
            .Build(new Asn1SignatureFactory("SHA256WITHRSA", _tsaKey.Private, Random), _tsaCertificate);
        var tokens = new TimeStampTokenGenerator(signer, Asn1DigestFactory.Get("SHA256"), new DerObjectIdentifier("1.3.6.1.4.1.59069.1.1"), true);
        tokens.SetCertificates(CollectionUtilities.CreateStore(new[] { _tsaCertificate, _rootCertificate }));
        var responses = new TimeStampResponseGenerator(tokens, TspAlgorithms.Allowed);
        return responses.Generate(new TimeStampRequest(request), BigInteger.ValueOf(Interlocked.Increment(ref _serial)), generationTime).GetEncoded();
    }

    private void Serve() {
        try {
            while (!_stop.IsCancellationRequested) {
                TcpClient client;
                try {
                    client = _listener.AcceptTcpClient();
                } catch (Exception) when (_stop.IsCancellationRequested) {
                    return;
                }
                lock (_clientGate) {
                    if (_stop.IsCancellationRequested) {
                        client.Dispose();
                        return;
                    }
                    _activeClient = client;
                }
                try {
                    Handle(client);
                } finally {
                    lock (_clientGate) {
                        _activeClient = null;
                    }
                }
            }
        } catch (Exception error) {
            _serverError ??= error;
        }
    }

    private void Handle(TcpClient client) {
        using (client) {
            try {
                NetworkStream stream;
                lock (_clientGate) {
                    if (_stop.IsCancellationRequested) {
                        return;
                    }
                    stream = client.GetStream();
                }
                (string method, string path, byte[] body) = ReadRequest(stream);
                byte[] response;
                string contentType;
                if (method == "POST" && path == "/ocsp") {
                    Interlocked.Increment(ref _ocspRequests);
                    response = CreateOcspResponse(body);
                    contentType = "application/ocsp-response";
                } else if (method == "GET" && path == "/root.crl") {
                    Interlocked.Increment(ref _crlRequests);
                    response = CreateCrl();
                    contentType = "application/pkix-crl";
                } else {
                    Write(stream, "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n", Array.Empty<byte>());
                    return;
                }
                Write(stream, "HTTP/1.1 200 OK\r\nContent-Type: " + contentType + "\r\nContent-Length: " + response.Length +
                    "\r\nConnection: close\r\n\r\n", response);
            } catch (IOException) {
            } catch (SocketException) {
            } catch (ObjectDisposedException) when (_stop.IsCancellationRequested) {
            } catch (Exception error) {
                _serverError ??= error;
            }
        }
    }

    private static (string Method, string Path, byte[] Body) ReadRequest(NetworkStream stream) {
        var header = new List<byte>();
        var buffer = new byte[1];
        while (header.Count < 65536) {
            if (stream.Read(buffer, 0, 1) == 0) break;
            header.Add(buffer[0]);
            int count = header.Count;
            if (count >= 4 && header[count - 4] == '\r' && header[count - 3] == '\n' && header[count - 2] == '\r' && header[count - 1] == '\n') break;
        }
        string[] lines = Encoding.ASCII.GetString(header.ToArray()).Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
        string[] requestLine = lines[0].Split(' ');
        int length = 0;
        foreach (string line in lines.Skip(1)) {
            if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)) {
                length = int.Parse(line.Substring("Content-Length:".Length).Trim(), System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        var body = new byte[length];
        int read = 0;
        while (read < length) {
            int chunk = stream.Read(body, read, length - read);
            if (chunk == 0) break;
            read += chunk;
        }
        return (requestLine[0], requestLine[1], body);
    }

    private static void Write(NetworkStream stream, string head, byte[] body) {
        byte[] headBytes = Encoding.ASCII.GetBytes(head);
        stream.Write(headBytes, 0, headBytes.Length);
        stream.Write(body, 0, body.Length);
        stream.Flush();
    }

    private static AsymmetricCipherKeyPair CreateKey() {
        var generator = new RsaKeyPairGenerator();
        generator.Init(new KeyGenerationParameters(Random, 2048));
        return generator.GenerateKeyPair();
    }

    private X509Certificate2 ToPlatform(BcCertificate certificate, AsymmetricCipherKeyPair key) {
        Pkcs12Store store = new Pkcs12StoreBuilder().Build();
        store.SetKeyEntry("key", new AsymmetricKeyEntry(key.Private),
            new[] { new X509CertificateEntry(certificate), new X509CertificateEntry(_rootCertificate) });
        using var output = new MemoryStream();
        store.Save(output, "test".ToCharArray(), Random);
#if NET9_0_OR_GREATER
        return X509CertificateLoader.LoadPkcs12(output.ToArray(), "test", X509KeyStorageFlags.Exportable | X509KeyStorageFlags.EphemeralKeySet);
#else
        return new X509Certificate2(output.ToArray(), "test", X509KeyStorageFlags.Exportable | X509KeyStorageFlags.EphemeralKeySet);
#endif
    }

    internal static X509Certificate2 Load(byte[] encoded) {
#if NET9_0_OR_GREATER
        return X509CertificateLoader.LoadCertificate(encoded);
#else
        return new X509Certificate2(encoded);
#endif
    }

    private sealed class DocumentTimestampSigner : IPdfExternalSigner {
        private readonly Rfc3161TimestampClient _client;

        internal DocumentTimestampSigner(Rfc3161TimestampClient client) => _client = client;

        public string Name => "OfficeIMO test document time-stamp";

        public byte[] Sign(PdfExternalSignatureRequest request) => _client.GetTimestamp(request.SignedContent);
    }
}
