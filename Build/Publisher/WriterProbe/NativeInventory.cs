using System.Security.Cryptography;
using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Bounded inventories of native bytes; these reports do not execute Publisher or active content.</summary>
internal static class NativeInventory {
    private const int MaximumBytes = 64 * 1024 * 1024;

    internal static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    internal static object Read(string path) {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (input.Length > MaximumBytes) throw new InvalidDataException("Probe input exceeds 64 MiB.");
        using var captured = new MemoryStream();
        input.CopyTo(captured);
        byte[] bytes = captured.ToArray();
        if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error))
            throw new InvalidDataException(error);
        return new {
            schema = "officeimo.publisher.writer-probe.inventory.v1",
            input = Path.GetFullPath(path),
            byteLength = bytes.Length,
            sha256 = Hash(bytes),
            rootClassId = compound!.RootEntry.ClassId,
            entries = compound.Entries.Select(entry => new {
                entry.Path, entry.ObjectType, entry.ClassId, entry.StateBits,
                entry.CreationTime, entry.ModifiedTime, entry.DirectoryOrder, entry.IsFallback
            }),
            streams = compound.Streams.OrderBy(entry => entry.Key, StringComparer.Ordinal).Select(entry => new {
                path = entry.Key, bytes = entry.Value.Length, sha256 = Hash(entry.Value),
                prefix = Convert.ToHexString(entry.Value.AsSpan(0, Math.Min(64, entry.Value.Length))).ToLowerInvariant()
            }),
            oleIdentification = compound.Streams.Where(entry => entry.Key.EndsWith("\u0001CompObj", StringComparison.OrdinalIgnoreCase))
                .Select(entry => new { path = entry.Key, identification = Identification(entry.Value) }).ToArray(),
            contents = InspectContents(compound),
            quill = InspectQuill(compound),
            officeArt = InspectOfficeArt(compound),
            officeImoRead = InspectManagedRead(bytes),
            nativeAcceptance = "not-run"
        };
    }

    private static object Identification(byte[] bytes) {
        try {
            if (bytes.Length < 28) throw new InvalidDataException("Truncated OLE CompObj header.");
            int offset = 28;
            string? ReadAnsi() {
                if (offset > bytes.Length - 4) throw new InvalidDataException("Truncated OLE identification length.");
                uint count = BitConverter.ToUInt32(bytes, offset); offset += 4;
                if (count == uint.MaxValue) {
                    if (offset > bytes.Length - 4) throw new InvalidDataException("Truncated numeric clipboard marker.");
                    offset += 4; return null;
                }
                if (count == 0) return null;
                if (count > 4096 || count > bytes.Length - offset || bytes[offset + (int)count - 1] != 0)
                    throw new InvalidDataException("Invalid bounded OLE identification string.");
                string value = System.Text.Encoding.ASCII.GetString(bytes, offset, (int)count - 1);
                offset += (int)count;
                return value;
            }
            string? userType = ReadAnsi(), clipboardFormat = ReadAnsi(), programId = ReadAnsi();
            return new {
                status = "decoded-identification",
                classId = new Guid(bytes.AsSpan(12, 16)), userType, clipboardFormat, programId,
                canonicalClipboardFormat = OfficeOleCompoundObjectReader.ClipboardFormat(bytes)
            };
        } catch (InvalidDataException exception) {
            return new { status = "unresolved", error = exception.Message };
        }
    }

    private static object InspectContents(OfficeCompoundFile compound) {
        if (!compound.Streams.TryGetValue("Contents", out byte[]? bytes)) return new { status = "missing" };
        try {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var reader = new PublisherBlockReader(data, context);
            int trailer = data.Offset(data.U32(0x1A));
            var trailerFields = reader.Chunk(trailer);
            var native = new PublisherContentsReader(data, context).Read();
            object headerProperties;
            try {
                int headerEnd = data.Offset(data.U32(0x1E));
                headerProperties = headerEnd >= 48 && headerEnd == 44 + data.U32(44)
                    ? new { status = "decoded-known-envelopes", end = headerEnd, blocks = Blocks(data, reader, reader.Chunk(44, headerEnd), 0) }
                    : (object)new { status = "not-mapped" };
            } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or OverflowException) {
                headerProperties = new { status = "unresolved", error = exception.Message };
            }
            return new {
                status = "decoded-known-records", native.Width, native.Height, trailer,
                header = Convert.ToHexString(bytes.AsSpan(0, Math.Min(32, bytes.Length))).ToLowerInvariant(),
                headerProperties,
                trailerFields = Blocks(data, reader, trailerFields, 0),
                chunks = native.Chunks.Values.OrderBy(chunk => chunk.Id).Select(chunk => new {
                    chunk.Id, chunk.Kind, chunk.Offset, chunk.End, chunk.Parent,
                    sha256 = Hash(bytes.AsSpan(chunk.Offset, chunk.End - chunk.Offset).ToArray()),
                    prefix = Convert.ToHexString(bytes.AsSpan(chunk.Offset, Math.Min(64, chunk.End - chunk.Offset))).ToLowerInvariant(),
                    fields = ChunkFields(data, reader, chunk)
                }).ToArray(),
                limitation = "Only block encodings understood by the existing decoder are interpreted. Unknown meanings remain unknown."
            };
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or OverflowException) {
            return new { status = "unresolved", error = exception.Message };
        }
    }

    private static object[] Blocks(PublisherBinaryData data, PublisherBlockReader reader, IReadOnlyList<PublisherBlock> fields, int depth) {
        if (depth > 16) return new object[] { new { status = "depth-limit" } };
        return fields.Select(field => (object)new {
            field.Id, field.Type, field.Offset, field.Length, field.Value,
            text = field.Type == 0xC0 ? data.Utf16(field.Offset + 4, field.Length - 4).TrimEnd('\0') : null,
            rawValue = field.Type is 0x28 or 0x38 or 0x48
                ? Convert.ToHexString(data.Bytes.AsSpan(field.Offset, field.Length)).ToLowerInvariant() : null,
            children = field.Type is 0x80 or 0x82 or 0x88 or 0x8A or 0x90 or 0x98 or 0xA0
                ? Children(data, reader, field, depth + 1) : null
        }).ToArray();
    }

    private static object ChunkFields(PublisherBinaryData data, PublisherBlockReader reader, PublisherContentChunk chunk) {
        try { return new { status = "decoded-known-envelopes", blocks = Blocks(data, reader, reader.Chunk(chunk.Offset, chunk.End), 0) }; }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or OverflowException) {
            return new { status = "unresolved", error = exception.Message };
        }
    }

    private static object Children(PublisherBinaryData data, PublisherBlockReader reader, PublisherBlock field, int depth) {
        try { return new { status = "decoded-known-envelopes", blocks = Blocks(data, reader, reader.Children(field), depth) }; }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or OverflowException) {
            return new { status = "opaque-payload", error = exception.Message };
        }
    }

    private static object InspectQuill(OfficeCompoundFile compound) {
        if (!compound.Streams.TryGetValue("Quill/QuillSub/CONTENTS", out byte[]? bytes)) return new { status = "missing" };
        try {
            var data = new PublisherBinaryData(bytes, "Quill/QuillSub/CONTENTS");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var chunks = PublisherQuillDirectory.Read(data, context);
            return new {
                status = "decoded-directory",
                chunks = chunks.Select(chunk => new {
                    chunk.Tag, chunk.Id, chunk.Offset, chunk.Length,
                    sha256 = Hash(bytes.AsSpan(chunk.Offset, chunk.Length).ToArray()),
                    prefix = Convert.ToHexString(bytes.AsSpan(chunk.Offset, Math.Min(96, chunk.Length))).ToLowerInvariant()
                }).ToArray()
            };
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or OverflowException) {
            return new { status = "unresolved", error = exception.Message };
        }
    }

    private static object InspectOfficeArt(OfficeCompoundFile compound) {
        if (!compound.Streams.TryGetValue("Escher/EscherStm", out byte[]? bytes)) return new { status = "missing" };
        try {
            var data = new PublisherBinaryData(bytes, "Escher/EscherStm");
            int budget = 100_000;
            return new { status = "decoded-record-envelopes", records = OfficeArtRecords(data, 0, bytes.Length, 0, ref budget) };
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException) {
            return new { status = "unresolved", error = exception.Message };
        }
    }

    private static object[] OfficeArtRecords(PublisherBinaryData data, int start, int end, int depth, ref int budget) {
        if (depth > 64) throw new InvalidDataException("OfficeArt inventory nesting exceeds 64.");
        var result = new List<object>();
        for (int position = start; position < end;) {
            if (--budget < 0) throw new InvalidDataException("OfficeArt inventory exceeds 100,000 records.");
            data.Range(position, 8, end);
            ushort initial = data.U16(position), kind = data.U16(position + 2);
            int length = data.Offset(data.U32(position + 4)), body = position + 8;
            data.Range(body, length, end);
            result.Add(new {
                offset = position, initial, kind, length,
                separator = kind is 0xF000 or 0xF002 && body + length < end ? (uint?)data.U32(body + length) : null,
                prefix = Convert.ToHexString(data.Bytes.AsSpan(body, Math.Min(64, length))).ToLowerInvariant(),
                children = (initial & 15) == 15 ? OfficeArtRecords(data, body, body + length, depth + 1, ref budget) : null
            });
            position = body + length;
            if (kind is 0xF000 or 0xF002 && position < end) { data.Range(position, 4, end); position += 4; }
        }
        return result.ToArray();
    }

    private static object InspectManagedRead(byte[] bytes) {
        try {
            PublisherDocument document = PublisherDocument.Load(bytes);
            return new {
                status = "accepted",
                pages = document.Pages.Select(page => new { page.Id, page.Width, page.Height, textFrames = page.TextFrames.Count, tables = page.Tables.Count }).ToArray(),
                stories = document.TextStories.Select(story => new { story.Id, story.Text }).ToArray(),
                images = document.Images.Count,
                diagnostics = document.ReadReport.FidelityDiagnostics.Select(diagnostic => new { diagnostic.Code, diagnostic.LossKind, diagnostic.Message }).ToArray()
            };
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or OverflowException) {
            return new { status = "rejected", error = exception.Message };
        }
    }
}
