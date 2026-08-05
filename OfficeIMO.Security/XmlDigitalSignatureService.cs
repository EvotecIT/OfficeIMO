#nullable enable
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;

namespace OfficeIMO.Security;

/// <summary>Bounded XML Digital Signature primitives used by format-owned signing workflows.</summary>
internal static class XmlDigitalSignatureService {
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The provider selects a closed XML DSig algorithm set and does not resolve caller-supplied transform implementations.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The provider rejects XSLT and limits XML DSig to statically referenced canonicalization and enveloped-signature transforms.")]
    internal static byte[] Create(XmlDigitalSignatureCreationRequest request) {
#if NETSTANDARD2_0 || NET472
        if (request == null) throw new ArgumentNullException(nameof(request));
#else
        ArgumentNullException.ThrowIfNull(request);
#endif
        EnsureXmlDsigAlgorithmRoots();
        ValidateCreationRequest(request);

        XmlDocument objectDocument = LoadXml(request.ObjectXml, request.MaxObjectBytes);
        using RSA? signingKey = request.SigningCertificate.GetRSAPrivateKey();
        if (signingKey == null) {
            throw new CryptographicException("XML signature creation requires an RSA certificate with an accessible private key.");
        }

        var packageObject = new DataObject {
            Id = request.ObjectId,
            Data = objectDocument.DocumentElement!.ChildNodes
        };
        XmlDocument signatureDocument = CreateXmlDocument();
        var signedXml = new SignedXml(signatureDocument) { SigningKey = signingKey };
        signedXml.Signature.Id = request.SignatureId;
        signedXml.SignedInfo!.CanonicalizationMethod = request.CanonicalizationMethod;
        signedXml.SignedInfo.SignatureMethod = request.SignatureMethod;
        signedXml.AddObject(packageObject);
        signedXml.AddReference(new Reference("#" + request.ObjectId) {
            Type = request.ObjectReferenceType,
            DigestMethod = request.DigestMethod
        });

        var keyInfo = new KeyInfo();
        var x509Data = new KeyInfoX509Data(request.SigningCertificate);
        if (request.AdditionalCertificates != null) {
            foreach (X509Certificate2 certificate in request.AdditionalCertificates) {
                if (certificate != null) x509Data.AddCertificate(certificate);
            }
        }
        keyInfo.AddClause(x509Data);
        signedXml.KeyInfo = keyInfo;
        signedXml.ComputeSignature();

        XmlElement signature = signedXml.GetXml();
        signatureDocument.AppendChild(signatureDocument.ImportNode(signature, deep: true));
        using var output = new BoundedMemoryStream(request.MaxOutputBytes);
        using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = false
        })) {
            signatureDocument.Save(writer);
        }
        return output.ToArray();
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The provider accepts only explicitly permitted XML DSig algorithms whose implementations are referenced directly.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The provider rejects XSLT and does not compile dynamic transforms.")]
    internal static XmlDigitalSignatureVerificationResult Verify(XmlDigitalSignatureVerificationRequest request) {
#if NETSTANDARD2_0 || NET472
        if (request == null) throw new ArgumentNullException(nameof(request));
#else
        ArgumentNullException.ThrowIfNull(request);
#endif
        EnsureXmlDsigAlgorithmRoots();
        ValidateVerificationRequest(request);
        var findings = new List<SecurityFinding>();
        XmlDocument document;
        try {
            document = LoadXml(request.SignatureXml, request.MaxSignatureBytes);
        } catch (Exception exception) when (exception is XmlException or IOException or InvalidDataException) {
            findings.Add(Finding(SecurityFindingSeverity.Error, "XmlSignatureMalformed",
                "The XML signature document is invalid: " + exception.Message));
            return Result(SecurityValidationStatus.Invalid, Array.Empty<X509Certificate2>(), findings);
        }

        XmlElement? signatureElement = document.DocumentElement;
        if (signatureElement == null ||
            signatureElement.LocalName != "Signature" ||
            signatureElement.NamespaceURI != XmlDigitalSignatureAlgorithms.Namespace) {
            findings.Add(Finding(SecurityFindingSeverity.Error, "XmlSignatureMalformed",
                "The document does not contain an XML DSig Signature root element."));
            return Result(SecurityValidationStatus.Invalid, Array.Empty<X509Certificate2>(), findings);
        }

        XmlElement? signedInfoElement = signatureElement.ChildNodes
            .OfType<XmlElement>()
            .FirstOrDefault(element =>
                element.LocalName == "SignedInfo" &&
                element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
        XmlElement? canonicalizationElement = signedInfoElement?.ChildNodes
            .OfType<XmlElement>()
            .FirstOrDefault(element =>
                element.LocalName == "CanonicalizationMethod" &&
                element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
        string canonicalizationMethod = canonicalizationElement?.GetAttribute("Algorithm") ?? string.Empty;
        if (!IsSupportedCanonicalizationMethod(canonicalizationMethod) ||
            !request.AllowedCanonicalizationMethods.Contains(canonicalizationMethod, StringComparer.Ordinal)) {
            findings.Add(Finding(SecurityFindingSeverity.Warning, "UnsupportedSignedInfoCanonicalizationMethod",
                "SignedInfo canonicalization method '" + canonicalizationMethod + "' is not accepted by provider and caller policy."));
            return Result(SecurityValidationStatus.Indeterminate, Array.Empty<X509Certificate2>(), findings);
        }

        XmlElement? signatureMethodElement = signedInfoElement?.ChildNodes
            .OfType<XmlElement>()
            .FirstOrDefault(element =>
                element.LocalName == "SignatureMethod" &&
                element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
        string signatureMethod = signatureMethodElement?.GetAttribute("Algorithm") ?? string.Empty;
        if (!IsSupportedSignatureMethod(signatureMethod) ||
            !request.AllowedSignatureMethods.Contains(signatureMethod, StringComparer.Ordinal)) {
            findings.Add(Finding(SecurityFindingSeverity.Warning, "UnsupportedSignatureMethod",
                "XML signature method '" + signatureMethod + "' is not accepted by provider and caller policy."));
            return Result(SecurityValidationStatus.Indeterminate, Array.Empty<X509Certificate2>(), findings);
        }

        XmlElement[] references = signedInfoElement?.ChildNodes
            .OfType<XmlElement>()
            .Where(element =>
                element.LocalName == "Reference" &&
                element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
            .ToArray() ?? Array.Empty<XmlElement>();
        if (references.Length > request.MaxReferences) {
            throw new InvalidDataException(
                "The XML signature contains more than " + request.MaxReferences + " SignedInfo references.");
        }
        foreach (XmlElement reference in references) {
            XmlElement? digestMethodElement = reference.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "DigestMethod" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
            string digestMethod = digestMethodElement?.GetAttribute("Algorithm") ?? string.Empty;
            if (!IsSupportedDigestMethod(digestMethod) ||
                !request.AllowedDigestMethods.Contains(digestMethod, StringComparer.Ordinal)) {
                findings.Add(Finding(SecurityFindingSeverity.Warning, "UnsupportedDigestMethod",
                    "XML reference digest method '" + digestMethod + "' is not accepted by provider and caller policy."));
                return Result(SecurityValidationStatus.Indeterminate, Array.Empty<X509Certificate2>(), findings);
            }
            foreach (XmlElement transform in reference.ChildNodes
                         .OfType<XmlElement>()
                         .Where(element =>
                             element.LocalName == "Transforms" &&
                             element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                         .SelectMany(element => element.ChildNodes
                             .OfType<XmlElement>()
                             .Where(child =>
                                 child.LocalName == "Transform" &&
                                 child.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace))) {
                string algorithm = transform.GetAttribute("Algorithm");
                if (IsSupportedReferenceTransform(algorithm) &&
                    request.AllowedReferenceTransforms.Contains(algorithm, StringComparer.Ordinal)) continue;
                findings.Add(Finding(SecurityFindingSeverity.Warning, "UnsupportedSignedInfoTransform",
                    "SignedInfo reference transform '" + algorithm + "' is not accepted by provider and caller policy."));
                return Result(SecurityValidationStatus.Indeterminate, Array.Empty<X509Certificate2>(), findings);
            }
        }

        var signedXml = new SignedXml(document) { Resolver = null! };
        try {
            signedXml.LoadXml(signatureElement);
        } catch (CryptographicException exception) {
            findings.Add(Finding(SecurityFindingSeverity.Error, "XmlSignatureMalformed",
                "The XML DSig structure is invalid: " + exception.Message));
            return Result(SecurityValidationStatus.Invalid, Array.Empty<X509Certificate2>(), findings);
        }

        foreach (object item in signedXml.SignedInfo!.References) {
            if (item is not Reference reference) continue;
            string uri = reference.Uri ?? string.Empty;
            if (uri.Length == 0 || uri[0] == '#') continue;
            findings.Add(Finding(SecurityFindingSeverity.Warning, "ExternalSignedInfoReference",
                "SignedInfo reference '" + uri + "' is not a local fragment and was not dereferenced."));
            return Result(SecurityValidationStatus.Indeterminate, Array.Empty<X509Certificate2>(), findings);
        }

        _ = XmlDigitalSignatureReferenceWorkCalculator.Measure(
            document,
            signedInfoElement,
            request.CertificateCandidates.Count,
            request.MaxTotalDigestWorkBytes);

        var matches = new List<X509Certificate2>();
        foreach (X509Certificate2 candidate in request.CertificateCandidates) {
            if (candidate == null) continue;
            try {
                using AsymmetricAlgorithm? publicKey = GetPublicKey(candidate);
                if (publicKey != null && signedXml.CheckSignature(publicKey)) matches.Add(candidate);
            } catch (CryptographicException) {
                // Continue through the bounded caller-supplied candidate set.
            }
        }

        if (matches.Count == 0) {
            findings.Add(Finding(SecurityFindingSeverity.Error, "XmlSignatureInvalid",
                "XML DSig signature-value or signed-object validation failed for every supplied certificate."));
            return Result(SecurityValidationStatus.Invalid, matches, findings);
        }
        findings.Add(Finding(SecurityFindingSeverity.Info, "XmlSignatureValid",
            "XML DSig signature-value and signed-object validation passed."));
        return Result(SecurityValidationStatus.Valid, matches, findings);
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Canonicalization is selected from a closed set of statically referenced XML DSig transforms.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Canonicalization does not use the XSLT transform or dynamic code generation.")]
    internal static byte[] Canonicalize(
        byte[] xml,
        string algorithm,
        string? inclusiveNamespacesPrefixList,
        long maxOutputBytes) {
#if NETSTANDARD2_0 || NET472
        if (xml == null) throw new ArgumentNullException(nameof(xml));
#else
        ArgumentNullException.ThrowIfNull(xml);
#endif
        if (string.IsNullOrWhiteSpace(algorithm)) throw new ArgumentException("A canonicalization algorithm is required.", nameof(algorithm));
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxOutputBytes);
#else
        if (maxOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
#endif
        XmlDocument document = LoadXml(xml, maxOutputBytes);
        Transform transform;
        switch (algorithm) {
            case XmlDigitalSignatureAlgorithms.CanonicalXml:
                transform = new XmlDsigC14NTransform();
                break;
            case XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments:
                transform = new XmlDsigC14NWithCommentsTransform();
                break;
            case XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml:
                transform = new XmlDsigExcC14NTransform {
                    InclusiveNamespacesPrefixList = inclusiveNamespacesPrefixList
                };
                break;
            case XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXmlWithComments:
                transform = new XmlDsigExcC14NTransform(includeComments: true) {
                    InclusiveNamespacesPrefixList = inclusiveNamespacesPrefixList
                };
                break;
            default:
                throw new NotSupportedException("XML canonicalization method " + algorithm + " is not supported.");
        }
        transform.LoadInput(document);
        using Stream canonical = (Stream)transform.GetOutput(typeof(Stream));
        using var output = new BoundedMemoryStream(maxOutputBytes);
        CopyBounded(canonical, output, maxOutputBytes);
        return output.ToArray();
    }

    private static void ValidateCreationRequest(XmlDigitalSignatureCreationRequest request) {
        if (request.MaxObjectBytes <= 0) throw new ArgumentOutOfRangeException(nameof(request), "MaxObjectBytes must be positive.");
        if (request.MaxOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(request), "MaxOutputBytes must be positive.");
        if (request.ObjectXml.LongLength > request.MaxObjectBytes) {
            throw new InvalidDataException("The XML signature object exceeds the configured byte limit.");
        }
        if (!request.SigningCertificate.HasPrivateKey) {
            throw new CryptographicException("The XML signing certificate must include a private key.");
        }
        if (!IsSupportedCanonicalizationMethod(request.CanonicalizationMethod)) {
            throw new NotSupportedException("XML canonicalization method " + request.CanonicalizationMethod + " is not supported.");
        }
        if (!IsSupportedSignatureMethod(request.SignatureMethod)) {
            throw new NotSupportedException("XML signature method " + request.SignatureMethod + " is not supported.");
        }
        if (!IsSupportedDigestMethod(request.DigestMethod)) {
            throw new NotSupportedException("XML digest method " + request.DigestMethod + " is not supported.");
        }
    }

    private static void ValidateVerificationRequest(XmlDigitalSignatureVerificationRequest request) {
        if (request.MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(request), "MaxSignatureBytes must be positive.");
        if (request.MaxReferences <= 0) throw new ArgumentOutOfRangeException(nameof(request), "MaxReferences must be positive.");
        if (request.MaxTotalDigestWorkBytes <= 0) throw new ArgumentOutOfRangeException(nameof(request), "MaxTotalDigestWorkBytes must be positive.");
        if (request.SignatureXml.LongLength > request.MaxSignatureBytes) {
            throw new InvalidDataException("The XML signature exceeds the configured byte limit.");
        }
        if (request.AllowedCanonicalizationMethods == null) throw new ArgumentException("AllowedCanonicalizationMethods cannot be null.", nameof(request));
        if (request.AllowedReferenceTransforms == null) throw new ArgumentException("AllowedReferenceTransforms cannot be null.", nameof(request));
        if (request.AllowedSignatureMethods == null) throw new ArgumentException("AllowedSignatureMethods cannot be null.", nameof(request));
        if (request.AllowedDigestMethods == null) throw new ArgumentException("AllowedDigestMethods cannot be null.", nameof(request));
    }

    private static bool IsSupportedCanonicalizationMethod(string algorithm) => algorithm is
        XmlDigitalSignatureAlgorithms.CanonicalXml or
        XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments or
        XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml or
        XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXmlWithComments;

    private static bool IsSupportedReferenceTransform(string algorithm) =>
        IsSupportedCanonicalizationMethod(algorithm) ||
        algorithm == XmlDigitalSignatureAlgorithms.EnvelopedSignatureTransform;

    private static bool IsSupportedSignatureMethod(string algorithm) => algorithm is
        XmlDigitalSignatureAlgorithms.RsaSha1 or
        XmlDigitalSignatureAlgorithms.RsaSha256 or
        XmlDigitalSignatureAlgorithms.RsaSha384 or
        XmlDigitalSignatureAlgorithms.RsaSha512;

    private static bool IsSupportedDigestMethod(string algorithm) => algorithm is
        XmlDigitalSignatureAlgorithms.Sha1 or
        XmlDigitalSignatureAlgorithms.Sha256 or
        XmlDigitalSignatureAlgorithms.Sha384 or
        XmlDigitalSignatureAlgorithms.Sha512;

    private static XmlDocument LoadXml(byte[] xml, long maxBytes) {
        if (xml.LongLength > maxBytes) throw new InvalidDataException("The XML document exceeds the configured byte limit.");
        using var input = new MemoryStream(xml, writable: false);
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = maxBytes
        });
        XmlDocument document = CreateXmlDocument();
        document.Load(reader);
        return document;
    }

    private static XmlDocument CreateXmlDocument() => new() {
        PreserveWhitespace = true,
        XmlResolver = null
    };

#if NET8_0_OR_GREATER
#pragma warning disable SYSLIB0021
    // SignedXml resolves algorithm URIs through CryptoConfig. The runtime cannot infer those dynamic
    // activations during trimming or NativeAOT, so the closed algorithm set accepted by this provider
    // must root the corresponding implementations in the owning library rather than in each application.
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor, typeof(SHA1Managed))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor, typeof(SHA256Managed))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor, typeof(SHA384Managed))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor, typeof(SHA512Managed))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(RSAPKCS1SignatureFormatter))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(RSAPKCS1SignatureDeformatter))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(XmlDsigC14NTransform))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(XmlDsigC14NWithCommentsTransform))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(XmlDsigExcC14NTransform))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(XmlDsigExcC14NWithCommentsTransform))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicMethods, typeof(XmlDsigEnvelopedSignatureTransform))]
    private static void EnsureXmlDsigAlgorithmRoots() { }
#pragma warning restore SYSLIB0021
#else
    private static void EnsureXmlDsigAlgorithmRoots() { }
#endif

    private static AsymmetricAlgorithm? GetPublicKey(X509Certificate2 certificate) {
        AsymmetricAlgorithm? publicKey = certificate.GetRSAPublicKey();
        publicKey ??= certificate.GetECDsaPublicKey();
#if NETSTANDARD2_0 || NETFRAMEWORK
#pragma warning disable SYSLIB0027
        publicKey ??= certificate.PublicKey.Key;
#pragma warning restore SYSLIB0027
#else
        publicKey ??= certificate.GetDSAPublicKey();
#endif
        return publicKey;
    }

    private static void CopyBounded(Stream source, Stream destination, long maxBytes) {
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maxBytes) throw new InvalidDataException("Canonical XML exceeds the configured byte limit.");
            destination.Write(buffer, 0, read);
        }
    }

    private static SecurityFinding Finding(SecurityFindingSeverity severity, string code, string message) =>
        new(severity, code, message);

    private static XmlDigitalSignatureVerificationResult Result(
        SecurityValidationStatus status,
        IReadOnlyList<X509Certificate2> matchingCertificates,
        IReadOnlyList<SecurityFinding> findings) =>
        new(status, matchingCertificates, findings);
}
