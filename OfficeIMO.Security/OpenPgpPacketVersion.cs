using Org.BouncyCastle.Bcpg.OpenPgp;

namespace OfficeIMO.Security;

/// <summary>
/// Recognises packets from newer OpenPGP versions (RFC 9580 version 6 keys and signatures, version 2 SEIPD) that the
/// RFC 4880 implementation cannot read, so they are reported as unsupported rather than malformed.
/// </summary>
internal static class OpenPgpPacketVersion {
    internal static bool TryDescribeUnsupported(byte[] encoded, out string message) {
        message = string.Empty;
        byte[] data;
        try {
            using Stream decoded = PgpUtilities.GetDecoderStream(new MemoryStream(encoded, writable: false));
            using var buffer = new MemoryStream();
            var chunk = new byte[4096];
            int read;
            while ((read = decoded.Read(chunk, 0, chunk.Length)) > 0 && buffer.Length < 64) buffer.Write(chunk, 0, read);
            data = buffer.ToArray();
        } catch (Exception ex) when (OpenPgpService.IsParseFailure(ex)) {
            return false;
        }

        int offset = 0;
        while (TryReadHeader(data, ref offset, out int tag, out int length)) {
            if (offset >= data.Length) return false;
            int version = data[offset];
            if (tag == 10 && length >= 0) {
                offset += length; // Marker packet.
                continue;
            }

            if ((tag == 1 || tag == 2 || tag == 4 || tag == 5 || tag == 6 || tag == 7 || tag == 14) && version >= 5) {
                message = "OpenPGP version " + version.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " packets (RFC 9580) are not supported; only version 4 keys and signatures (RFC 4880) are.";
                return true;
            }

            if (tag == 18 && version == 2) {
                message = "Version 2 SEIPD (AEAD) encrypted data (RFC 9580) is not supported.";
                return true;
            }

            return false;
        }

        return false;
    }

    private static bool TryReadHeader(byte[] data, ref int offset, out int tag, out int length) {
        tag = 0;
        length = -1;
        if (offset >= data.Length || (data[offset] & 0x80) == 0) return false;
        byte first = data[offset++];
        if ((first & 0x40) != 0) {
            tag = first & 0x3F;
            if (offset >= data.Length) return false;
            byte octet = data[offset++];
            if (octet < 192) {
                length = octet;
            } else if (octet < 224) {
                if (offset >= data.Length) return false;
                length = ((octet - 192) << 8) + data[offset++] + 192;
            } else if (octet == 255) {
                if (offset + 4 > data.Length) return false;
                length = (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
                offset += 4;
            }

            return true;
        }

        tag = (first >> 2) & 0x0F;
        int lengthType = first & 0x03;
        int bytes = lengthType == 0 ? 1 : lengthType == 1 ? 2 : lengthType == 2 ? 4 : 0;
        if (offset + bytes > data.Length) return false;
        length = 0;
        for (int index = 0; index < bytes; index++) length = (length << 8) | data[offset++];
        if (bytes == 0) length = -1;
        return true;
    }
}
