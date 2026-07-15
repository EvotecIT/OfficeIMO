using System.Globalization;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Text;

#pragma warning disable CA1510 // Cross-target guard code supports netstandard2.0 and net472.

namespace OfficeIMO.Pdf.Cryptography;

internal sealed class PdfManagedCmsDocument {
    internal const string DataOid = "1.2.840.113549.1.7.1";
    internal const string SignedDataOid = "1.2.840.113549.1.7.2";
    internal const string TstInfoOid = "1.2.840.113549.1.9.16.1.4";
    internal const string MessageDigestOid = "1.2.840.113549.1.9.4";
    internal const string SigningTimeOid = "1.2.840.113549.1.9.5";
    internal const string SignatureTimestampOid = "1.2.840.113549.1.9.16.2.14";

    internal string ContentTypeOid { get; private set; } = string.Empty;
    internal byte[]? EncapsulatedContent { get; private set; }
    internal string DigestAlgorithmOid { get; private set; } = string.Empty;
    internal string SignatureAlgorithmOid { get; private set; } = string.Empty;
    internal byte[] SignatureValue { get; private set; } = Array.Empty<byte>();
    internal byte[]? SignedAttributes { get; private set; }
    internal byte[]? MessageDigest { get; private set; }
    internal DateTimeOffset? SigningTime { get; private set; }
    internal byte[]? SignerSerialNumber { get; private set; }
    internal byte[]? SignerIssuer { get; private set; }
    internal byte[]? SignerSubjectKeyIdentifier { get; private set; }
    internal List<X509Certificate2> Certificates { get; } = new List<X509Certificate2>();
    internal List<byte[]> SignatureTimestamps { get; } = new List<byte[]>();

    internal static PdfManagedCmsDocument Parse(byte[] encoded) {
        if (encoded == null) throw new ArgumentNullException(nameof(encoded));
        if (encoded.Length > 64 * 1024 * 1024) throw new InvalidDataException("CMS container exceeds the managed parser limit.");
        var root = new PdfDerReader(encoded);
        PdfDerElement contentInfo = root.Read(0x30);
        root.EnsureEnd();
        var contentInfoReader = contentInfo.Reader();
        string outerType = ReadOid(contentInfoReader.Read(0x06));
        if (!string.Equals(outerType, SignedDataOid, StringComparison.Ordinal)) throw new InvalidDataException("CMS ContentInfo is not SignedData.");
        PdfDerElement explicitContent = contentInfoReader.Read(0xA0);
        contentInfoReader.EnsureEnd();
        var explicitReader = explicitContent.Reader();
        PdfDerElement signedData = explicitReader.Read(0x30);
        explicitReader.EnsureEnd();

        var result = new PdfManagedCmsDocument();
        result.ReadSignedData(signedData);
        return result;
    }

    private void ReadSignedData(PdfDerElement signedData) {
        var reader = signedData.Reader();
        reader.Read(0x02);
        reader.Read(0x31);
        ReadEncapsulatedContentInfo(reader.Read(0x30));
        PdfDerElement next = reader.Read();
        if (next.Tag == 0xA0) {
            ReadCertificates(next);
            next = reader.Read();
        }
        if (next.Tag == 0xA1) next = reader.Read();
        if (next.Tag != 0x31) throw new InvalidDataException("CMS signerInfos set is missing.");
        var signers = next.Reader();
        if (!signers.HasData) throw new InvalidDataException("CMS does not contain a signer.");
        ReadSignerInfo(signers.Read(0x30));
        reader.EnsureEnd();
    }

    private void ReadEncapsulatedContentInfo(PdfDerElement value) {
        var reader = value.Reader();
        ContentTypeOid = ReadOid(reader.Read(0x06));
        if (reader.HasData) {
            PdfDerElement explicitContent = reader.Read(0xA0);
            var content = explicitContent.Reader();
            var segments = new List<byte[]>();
            while (content.HasData) segments.Add(content.Read(0x04).Content());
            EncapsulatedContent = PdfDerCodec.Concatenate(segments.ToArray());
        }
        reader.EnsureEnd();
    }

    private void ReadCertificates(PdfDerElement value) {
        var reader = value.Reader();
        while (reader.HasData) {
            PdfDerElement certificate = reader.Read();
            if (certificate.Tag != 0x30) continue;
            try {
#if NET9_0_OR_GREATER
                Certificates.Add(X509CertificateLoader.LoadCertificate(certificate.Encoded()));
#else
                Certificates.Add(new X509Certificate2(certificate.Encoded()));
#endif
            } catch (System.Security.Cryptography.CryptographicException) { }
        }
    }

    private void ReadSignerInfo(PdfDerElement value) {
        var reader = value.Reader();
        reader.Read(0x02);
        PdfDerElement sid = reader.Read();
        if (sid.Tag == 0x30) ReadIssuerAndSerial(sid);
        else if (sid.Tag == 0x80) SignerSubjectKeyIdentifier = sid.Content();
        DigestAlgorithmOid = ReadAlgorithmIdentifier(reader.Read(0x30));
        PdfDerElement next = reader.Read();
        if (next.Tag == 0xA0) {
            SignedAttributes = PdfDerCodec.ReplaceTag(next.Encoded(), 0x31);
            ReadAttributes(next, signed: true);
            next = reader.Read(0x30);
        } else if (next.Tag != 0x30) {
            throw new InvalidDataException("CMS signature algorithm is missing.");
        }
        SignatureAlgorithmOid = ReadAlgorithmIdentifier(next);
        SignatureValue = reader.Read(0x04).Content();
        if (reader.HasData) {
            PdfDerElement unsignedAttributes = reader.Read();
            if (unsignedAttributes.Tag == 0xA1) ReadAttributes(unsignedAttributes, signed: false);
        }
        reader.EnsureEnd();
    }

    private void ReadIssuerAndSerial(PdfDerElement value) {
        var reader = value.Reader();
        SignerIssuer = reader.Read(0x30).Encoded();
        SignerSerialNumber = NormalizeInteger(reader.Read(0x02).Content());
    }

    private void ReadAttributes(PdfDerElement value, bool signed) {
        var reader = value.Reader();
        while (reader.HasData) {
            PdfDerElement attribute = reader.Read(0x30);
            var attributeReader = attribute.Reader();
            string oid = ReadOid(attributeReader.Read(0x06));
            PdfDerElement values = attributeReader.Read(0x31);
            if (signed && string.Equals(oid, MessageDigestOid, StringComparison.Ordinal)) {
                var valueReader = values.Reader();
                if (valueReader.HasData) MessageDigest = valueReader.Read(0x04).Content();
            } else if (signed && string.Equals(oid, SigningTimeOid, StringComparison.Ordinal)) {
                var valueReader = values.Reader();
                if (valueReader.HasData) SigningTime = ReadTime(valueReader.Read());
            } else if (!signed && string.Equals(oid, SignatureTimestampOid, StringComparison.Ordinal)) {
                var valueReader = values.Reader();
                while (valueReader.HasData) SignatureTimestamps.Add(valueReader.Read().Encoded());
            }
        }
    }

    internal X509Certificate2? FindSignerCertificate(X509Certificate2Collection extraCertificates) {
        foreach (X509Certificate2 certificate in Certificates) if (MatchesSigner(certificate)) return certificate;
        foreach (X509Certificate2 certificate in extraCertificates) if (MatchesSigner(certificate)) return certificate;
        return null;
    }

    private bool MatchesSigner(X509Certificate2 certificate) {
        if (SignerSubjectKeyIdentifier != null) {
            byte[]? certificateIdentifier = ReadSubjectKeyIdentifier(certificate);
            return certificateIdentifier != null && FixedTimeEquals(certificateIdentifier, SignerSubjectKeyIdentifier);
        }
        if (SignerSerialNumber == null || SignerIssuer == null) return false;
        byte[] serial = certificate.GetSerialNumber();
        Array.Reverse(serial);
        serial = NormalizeInteger(serial);
        return FixedTimeEquals(serial, SignerSerialNumber) && certificate.IssuerName.RawData.SequenceEqual(SignerIssuer);
    }

    private static byte[]? ReadSubjectKeyIdentifier(X509Certificate2 certificate) {
        foreach (X509Extension extension in certificate.Extensions) {
            if (!string.Equals(extension.Oid?.Value, "2.5.29.14", StringComparison.Ordinal)) continue;
            try {
                var reader = new PdfDerReader(extension.RawData);
                byte[] identifier = reader.Read(0x04).Content();
                reader.EnsureEnd();
                return identifier;
            } catch (InvalidDataException) {
                return null;
            }
        }
        return null;
    }

    internal static string ReadAlgorithmIdentifier(PdfDerElement value) {
        var reader = value.Reader();
        return ReadOid(reader.Read(0x06));
    }

    internal static string ReadOid(PdfDerElement value) {
        byte[] content = value.Content();
        if (content.Length == 0) throw new InvalidDataException("DER object identifier is empty.");
        var components = new List<uint>();
        int index = 0;
        uint combined = ReadBase128(content, ref index);
        uint first = combined < 40 ? 0U : combined < 80 ? 1U : 2U;
        components.Add(first);
        components.Add(combined - (first * 40U));
        while (index < content.Length) components.Add(ReadBase128(content, ref index));
        return string.Join(".", components.Select(static component => component.ToString(CultureInfo.InvariantCulture)));
    }

    internal static DateTimeOffset? ReadTime(PdfDerElement value) {
        string text = Encoding.ASCII.GetString(value.Content());
        if (value.Tag == 0x17) {
            return DateTimeOffset.TryParseExact(text, "yyMMddHHmmss'Z'", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTimeOffset utcTime)
                ? utcTime
                : null;
        }
        if (value.Tag != 0x18 || text.Length < 15 || text[text.Length - 1] != 'Z') return null;
        string valueWithoutZone = text.Substring(0, text.Length - 1);
        string wholeSeconds = valueWithoutZone;
        string fraction = string.Empty;
        int separator = valueWithoutZone.IndexOf('.');
        if (separator >= 0) {
            wholeSeconds = valueWithoutZone.Substring(0, separator);
            fraction = valueWithoutZone.Substring(separator + 1);
            if (fraction.Length == 0 || fraction.Any(static character => character < '0' || character > '9')) return null;
        }
        if (!DateTimeOffset.TryParseExact(wholeSeconds + "Z", "yyyyMMddHHmmss'Z'", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTimeOffset generalizedTime)) return null;
        if (fraction.Length == 0) return generalizedTime;
        string ticksText = fraction.Length >= 7 ? fraction.Substring(0, 7) : fraction.PadRight(7, '0');
        return long.TryParse(ticksText, NumberStyles.None, CultureInfo.InvariantCulture, out long ticks)
            ? generalizedTime.AddTicks(ticks)
            : null;
    }

    internal static byte[] NormalizeInteger(byte[] value) {
        int offset = 0;
        while (offset + 1 < value.Length && value[offset] == 0) offset++;
        var result = new byte[value.Length - offset];
        Buffer.BlockCopy(value, offset, result, 0, result.Length);
        return result;
    }

    internal static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int i = 0; i < left.Length; i++) difference |= left[i] ^ right[i];
        return difference == 0;
    }

    private static uint ReadBase128(byte[] value, ref int index) {
        uint result = 0;
        int count = 0;
        while (index < value.Length) {
            byte current = value[index++];
            if (++count > 5 || result > (uint.MaxValue >> 7)) throw new InvalidDataException("DER object identifier component is too large.");
            result = (result << 7) | (uint)(current & 0x7F);
            if ((current & 0x80) == 0) return result;
        }
        throw new InvalidDataException("DER object identifier is truncated.");
    }
}
