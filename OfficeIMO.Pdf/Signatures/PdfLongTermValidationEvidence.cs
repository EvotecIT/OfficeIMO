namespace OfficeIMO.Pdf;

/// <summary>DER-encoded validation material to append for one verified PDF signature.</summary>
public sealed class PdfLongTermValidationEvidence {
    private readonly byte[][] _certificates;
    private readonly byte[][] _ocspResponses;
    private readonly byte[][] _certificateRevocationLists;

    /// <summary>Creates validation evidence for a signature value object.</summary>
    /// <param name="signatureObjectNumber">Object number of the signature dictionary to enrich.</param>
    /// <param name="certificates">DER-encoded X.509 certificates used during validation.</param>
    /// <param name="ocspResponses">DER-encoded OCSPResponse values used during validation.</param>
    /// <param name="certificateRevocationLists">DER-encoded X.509 CRLs used during validation.</param>
    public PdfLongTermValidationEvidence(
        int signatureObjectNumber,
        IEnumerable<byte[]>? certificates = null,
        IEnumerable<byte[]>? ocspResponses = null,
        IEnumerable<byte[]>? certificateRevocationLists = null) {
        if (signatureObjectNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(signatureObjectNumber), signatureObjectNumber, "Signature object number must be positive.");
        }

        SignatureObjectNumber = signatureObjectNumber;
        _certificates = Snapshot(certificates, nameof(certificates));
        _ocspResponses = Snapshot(ocspResponses, nameof(ocspResponses));
        _certificateRevocationLists = Snapshot(certificateRevocationLists, nameof(certificateRevocationLists));
        if (_certificates.Length == 0 && _ocspResponses.Length == 0 && _certificateRevocationLists.Length == 0) {
            throw new ArgumentException("At least one certificate, OCSP response, or CRL is required.");
        }
    }

    /// <summary>Object number of the signature dictionary to enrich.</summary>
    public int SignatureObjectNumber { get; }

    /// <summary>DER-encoded X.509 certificates used during validation.</summary>
    public IReadOnlyList<byte[]> Certificates => Clone(_certificates);

    /// <summary>DER-encoded OCSPResponse values used during validation.</summary>
    public IReadOnlyList<byte[]> OcspResponses => Clone(_ocspResponses);

    /// <summary>DER-encoded X.509 CRLs used during validation.</summary>
    public IReadOnlyList<byte[]> CertificateRevocationLists => Clone(_certificateRevocationLists);

    /// <summary>True when OCSP or CRL status material is present.</summary>
    public bool HasRevocationEvidence => _ocspResponses.Length > 0 || _certificateRevocationLists.Length > 0;

    internal IReadOnlyList<byte[]> CertificateValues => _certificates;
    internal IReadOnlyList<byte[]> OcspValues => _ocspResponses;
    internal IReadOnlyList<byte[]> CrlValues => _certificateRevocationLists;

    private static byte[][] Snapshot(IEnumerable<byte[]>? values, string parameterName) {
        if (values is null) {
            return Array.Empty<byte[]>();
        }

        byte[][] result = values.Select(static value => value is null ? null! : (byte[])value.Clone()).ToArray();
        for (int i = 0; i < result.Length; i++) {
            if (result[i] is null) {
                throw new ArgumentException("Validation evidence cannot contain null values.", parameterName);
            }

            ValidateDerSequence(result[i], parameterName);
        }

        return result;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<byte[]> Clone(byte[][] values) =>
        Array.AsReadOnly(values.Select(static value => (byte[])value.Clone()).ToArray());

    private static void ValidateDerSequence(byte[] value, string parameterName) {
        if (value.Length < 2 || value[0] != 0x30) {
            throw new ArgumentException("Validation evidence must be a DER-encoded ASN.1 SEQUENCE.", parameterName);
        }

        int lengthByte = value[1];
        int headerLength;
        int contentLength;
        if ((lengthByte & 0x80) == 0) {
            headerLength = 2;
            contentLength = lengthByte;
        } else {
            int lengthBytes = lengthByte & 0x7F;
            if (lengthBytes == 0 || lengthBytes > 4 || value.Length < 2 + lengthBytes) {
                throw new ArgumentException("Validation evidence must use a finite DER length.", parameterName);
            }

            headerLength = 2 + lengthBytes;
            contentLength = 0;
            for (int i = 0; i < lengthBytes; i++) {
                contentLength = checked((contentLength << 8) | value[2 + i]);
            }
        }

        if (contentLength != value.Length - headerLength) {
            throw new ArgumentException("Validation evidence DER length does not match its byte count.", parameterName);
        }
    }
}
