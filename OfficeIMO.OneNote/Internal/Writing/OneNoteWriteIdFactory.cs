using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.OneNote;

/// <summary>Allocates native identifiers, optionally within an explicit reproducible artifact namespace.</summary>
internal sealed class OneNoteWriteIdFactory {
    private readonly byte[]? _prefix;
    private ulong _next;

    internal OneNoteWriteIdFactory(Guid? seed = null, string scope = "model") {
        if (!seed.HasValue) return;
        byte[] name = Encoding.UTF8.GetBytes(scope);
        _prefix = new byte[16 + name.Length];
        Buffer.BlockCopy(seed.Value.ToByteArray(), 0, _prefix, 0, 16);
        Buffer.BlockCopy(name, 0, _prefix, 16, name.Length);
    }

    internal Guid NewGuid() {
        if (_prefix == null) return Guid.NewGuid();
        byte[] input = new byte[_prefix.Length + 8];
        Buffer.BlockCopy(_prefix, 0, input, 0, _prefix.Length);
        ulong sequence = checked(++_next);
        for (int index = 0; index < 8; index++) input[_prefix.Length + index] = (byte)(sequence >> (8 * index));
        byte[] hash;
        using (SHA256 algorithm = SHA256.Create()) hash = algorithm.ComputeHash(input);
        byte[] bytes = new byte[16];
        Buffer.BlockCopy(hash, 0, bytes, 0, bytes.Length);
        bytes[7] = (byte)((bytes[7] & 0x0f) | 0x80);
        bytes[8] = (byte)((bytes[8] & 0x3f) | 0x80);
        return new Guid(bytes);
    }

    internal OneNoteExtendedGuid New() => new OneNoteExtendedGuid(NewGuid(), 1, 17);
}
