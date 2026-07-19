namespace OfficeIMO.Email.AddressBook;

internal static class OabBinary {
    internal static uint UInt32(byte[] buffer, int offset) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || offset > buffer.Length - 4) throw new InvalidDataException("Unexpected end of OAB data.");
        return unchecked((uint)(buffer[offset] |
            (buffer[offset + 1] << 8) |
            (buffer[offset + 2] << 16) |
            (buffer[offset + 3] << 24)));
    }

    internal static uint ReadUInt32(Stream stream, string location) {
        byte[] bytes = ReadExactly(stream, 4, location);
        return UInt32(bytes, 0);
    }

    internal static byte[] ReadExactly(Stream stream, int count, string location) {
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        var result = new byte[count];
        int offset = 0;
        while (offset < count) {
            int read = stream.Read(result, offset, count - offset);
            if (read <= 0) throw new InvalidDataException(string.Concat("Unexpected end of OAB data at ", location, "."));
            offset += read;
        }
        return result;
    }

    internal static void Seek(OabSource source, Stream stream, long relativeOffset, string location) {
        if (relativeOffset < 0 || relativeOffset > source.Length) {
            throw new InvalidDataException(string.Concat("OAB offset is outside the source at ", location, "."));
        }
        stream.Position = checked(source.BaseOffset + relativeOffset);
    }
}
