using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.OneNote;

/// <summary>Allocates native identifiers, optionally within an explicit reproducible artifact namespace.</summary>
internal sealed class OneNoteWriteIdFactory {
    private readonly byte[]? _prefix;
    private ulong _next;
    private readonly HashSet<Guid> _reserved = new HashSet<Guid>();

    internal OneNoteWriteIdFactory(Guid? seed = null, string scope = "model") {
        if (!seed.HasValue) return;
        byte[] name = Encoding.UTF8.GetBytes(scope);
        _prefix = new byte[16 + name.Length];
        Buffer.BlockCopy(seed.Value.ToByteArray(), 0, _prefix, 0, 16);
        Buffer.BlockCopy(name, 0, _prefix, 16, name.Length);
    }

    internal void Reserve(IEnumerable<Guid>? identities) {
        if (_prefix == null || identities == null) return;
        foreach (Guid identity in identities) if (identity != Guid.Empty) _reserved.Add(identity);
    }

    internal Guid NewGuid() {
        if (_prefix == null) return Guid.NewGuid();
        ulong sequence = checked(++_next);
        uint attempt = 0;
        while (true) {
            // A collision changes this slot's salt, never the traversal counter. Retained-ID
            // calls still consume one slot, so unchanged repeat writes keep synthetic IDs stable.
            byte[] input = new byte[_prefix.Length + 8 + (attempt == 0 ? 0 : 4)];
            Buffer.BlockCopy(_prefix, 0, input, 0, _prefix.Length);
            for (int index = 0; index < 8; index++) input[_prefix.Length + index] = (byte)(sequence >> (8 * index));
            if (attempt != 0) for (int index = 0; index < 4; index++) input[_prefix.Length + 8 + index] = (byte)(attempt >> (8 * index));
            byte[] hash;
            using (SHA256 algorithm = SHA256.Create()) hash = algorithm.ComputeHash(input);
            byte[] bytes = new byte[16];
            Buffer.BlockCopy(hash, 0, bytes, 0, bytes.Length);
            bytes[7] = (byte)((bytes[7] & 0x0f) | 0x80);
            bytes[8] = (byte)((bytes[8] & 0x3f) | 0x80);
            Guid generated = new Guid(bytes);
            if (_reserved.Add(generated)) return generated;
            attempt = checked(attempt + 1);
        }
    }

    internal OneNoteExtendedGuid New() => new OneNoteExtendedGuid(NewGuid(), 1, 17);
}
