using System.IO;
using Org.BouncyCastle.Asn1;

namespace OfficeIMO.Security;

/// <summary>Bounded ASN.1 container helpers shared by format-specific security adapters.</summary>
public static class SecurityEncoding {
    /// <summary>
    /// Decodes one ASN.1 object and returns its canonical encoding. Optional trailing zero padding is accepted for
    /// fixed-width containers such as a PDF signature <c>/Contents</c> value.
    /// </summary>
    public static byte[] NormalizeSingleAsn1Object(
        byte[] encoded,
        bool allowTrailingZeroPadding = false,
        long maxEncodedBytes = 512L * 1024 * 1024) {
#if NETSTANDARD2_0 || NET472
        if (encoded == null) throw new ArgumentNullException(nameof(encoded));
#else
        ArgumentNullException.ThrowIfNull(encoded);
#endif
        SecurityLimits.EnsureBufferWithinLimit(encoded, maxEncodedBytes, nameof(encoded));
        using var stream = new MemoryStream(encoded, writable: false);
        using var input = new Asn1InputStream(stream, checked((int)Math.Min(encoded.LongLength, int.MaxValue)));
        Asn1Object? value = input.ReadObject();
        if (value == null) throw new InvalidDataException("The input does not contain an ASN.1 object.");
        int next;
        while ((next = input.ReadByte()) >= 0) {
            if (!allowTrailingZeroPadding || next != 0) {
                throw new InvalidDataException("Unexpected bytes follow the first ASN.1 object.");
            }
        }
        return value.GetEncoded();
    }
}
