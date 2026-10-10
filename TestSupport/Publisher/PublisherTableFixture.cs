using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.TestSupport;

// Changes declared records in Sample.pub. These are codec fixtures, not
// independently authored Publisher documents or native-rendering evidence.
internal static class PublisherTableFixture {
    internal static byte[] MergeFirstRow(byte[] publication) => Rewrite(publication, streams => {
        byte[] contents = streams["Contents"];
        Block[] tableFields = Chunk(contents, ChunkOffset(contents, 299));
        uint cellsId = BitConverter.ToUInt32(contents, tableFields.Single(field => field.Id == 0x6B).Offset);
        Block[] cellFields = Chunk(contents, ChunkOffset(contents, cellsId));
        Block[] cells = Children(contents, cellFields.Single(field => field.Id == 2));
        Assert.Equal(6, cells.Length);
        // Retain the second cell's existing coordinate fields and widen it to
        // the first column; the first definition is now an ignored native slot.
        contents[cells[0].Offset - 2] = 0xFE;
        WriteScalar(contents, Children(contents, cells[1]).Single(field => field.Id == 3), 0);
        WriteScalar(contents, cellFields.Single(field => field.Id == 1), 5);
        byte[] quill = streams["Quill/QuillSub/CONTENTS"];
        (int descriptor, int data) = TableTextChunk(quill);
        Assert.Equal(5U, BitConverter.ToUInt32(quill, data));
        Write(quill, data, 4); // Native count is one less than the cell count.
        Buffer.BlockCopy(quill, data + 16, quill, data + 12, 5 * 4);
    });

    internal static byte[] WithoutTextMapping(byte[] publication) => Rewrite(publication, streams => {
        byte[] quill = streams["Quill/QuillSub/CONTENTS"];
        (int descriptor, int data) = TableTextChunk(quill);
        quill[descriptor + 4] = (byte)'X'; // TCD becomes TCX, leaving complete stories intact.
    });

    internal static byte[] MoveToMaster(byte[] publication, uint sourcePageId, uint masterPageId) =>
        Rewrite(publication, streams => {
            byte[] contents = streams["Contents"];
            Block sourceList = Chunk(contents, ChunkOffset(contents, sourcePageId)).Single(field => field.Id == 2);
            Block table = Children(contents, sourceList).Single(field => field.Type == 0x70 && BitConverter.ToUInt32(contents, field.Offset) == 299);
            byte[] reference = contents.Skip(table.Offset - 2).Take(table.Length + 2).ToArray();
            streams["Contents"] = RebuildContents(contents, new Dictionary<uint, byte[]> {
                [sourcePageId] = RewriteShapeList(contents, sourcePageId, reference, remove: true),
                [masterPageId] = RewriteShapeList(contents, masterPageId, reference, remove: false)
            });
        });

    internal static byte[] CopyToObject(byte[] publication, uint objectId) => Rewrite(publication, streams => {
        byte[] contents = streams["Contents"];
        int original = ChunkOffset(contents, 299);
        byte[] table = contents.Skip(original).Take(checked((int)BitConverter.ToUInt32(contents, original))).ToArray();
        int trailer = checked((int)BitConverter.ToUInt32(contents, 0x1A));
        Block directory = Chunk(contents, trailer).Single(field => field.Type == 0x90);
        Block slot = Children(contents, directory)[checked((int)objectId)];
        WriteScalar(contents, Children(contents, slot).Single(field => field.Id == 2), 0x10);
        streams["Contents"] = RebuildContents(contents, new Dictionary<uint, byte[]> { [objectId] = table });
    });

    internal static byte[] NamePage(byte[] publication, uint pageId, string name) => Rewrite(publication, streams => {
        byte[] contents = streams["Contents"];
        Block[] fields = Chunk(contents, ChunkOffset(contents, pageId));
        using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
        writer.Write(0);
        foreach (Block field in fields.Where(field => field.Id != 0x0E))
            writer.Write(contents, field.Offset - 2, field.Length + 2);
        byte[] text = Encoding.Unicode.GetBytes(name + '\0');
        writer.Write((byte)0x0E); writer.Write((byte)0xC0); writer.Write(4 + text.Length); writer.Write(text);
        byte[] page = output.ToArray(); Write(page, 0, checked((uint)page.Length));
        streams["Contents"] = RebuildContents(contents, new Dictionary<uint, byte[]> { [pageId] = page });
    });

    private static byte[] RewriteShapeList(byte[] bytes, uint pageId, byte[] reference, bool remove) {
        Block[] fields = Chunk(bytes, ChunkOffset(bytes, pageId));
        Block? list = fields.Where(field => field.Id == 2).Select(field => (Block?)field).SingleOrDefault();
        var children = list.HasValue ? Children(bytes, list.Value).Select(field => bytes.Skip(field.Offset - 2).Take(field.Length + 2).ToArray()).ToList()
            : new List<byte[]>();
        if (remove) Assert.True(children.RemoveAll(child => child.SequenceEqual(reference)) == 1);
        else children.Add(reference);
        using var listBytes = new MemoryStream(); using var listWriter = new BinaryWriter(listBytes);
        listWriter.Write((byte)2); listWriter.Write(list?.Type ?? (byte)0x90);
        listWriter.Write(4 + children.Sum(child => child.Length));
        foreach (byte[] child in children) listWriter.Write(child);
        using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
        writer.Write(0); bool replaced = false;
        foreach (Block field in fields) {
            if (field.Id == 2) { writer.Write(listBytes.ToArray()); replaced = true; }
            else writer.Write(bytes, field.Offset - 2, field.Length + 2);
        }
        if (!replaced) writer.Write(listBytes.ToArray());
        byte[] result = output.ToArray(); Write(result, 0, (uint)result.Length);
        return result;
    }

    private static byte[] RebuildContents(byte[] bytes, Dictionary<uint, byte[]> replacements) {
        int trailer = checked((int)BitConverter.ToUInt32(bytes, 0x1A));
        Block directory = Chunk(bytes, trailer).Single(field => field.Type == 0x90);
        var chunks = Children(bytes, directory).Select((slot, id) => (Id: (uint)id, Slot: slot))
            .Where(item => item.Slot.Type == 0x88)
            .Select(item => (item.Id, OffsetField: Children(bytes, item.Slot).Where(field => field.Id == 4).Select(field => (Block?)field).SingleOrDefault()))
            .Where(item => item.OffsetField.HasValue)
            .Select(item => (item.Id, Field: item.OffsetField!.Value, Offset: checked((int)BitConverter.ToUInt32(bytes, item.OffsetField.Value.Offset))))
            .OrderBy(item => item.Offset).ToArray();
        byte[] tail = bytes.Skip(trailer).ToArray();
        using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
        writer.Write(bytes, 0, chunks[0].Offset);
        for (int index = 0; index < chunks.Length; index++) {
            var chunk = chunks[index];
            WriteScalar(tail, chunk.Field with { Offset = chunk.Field.Offset - trailer }, checked((uint)output.Position));
            int boundary = index + 1 < chunks.Length ? chunks[index + 1].Offset : trailer;
            if (replacements.TryGetValue(chunk.Id, out byte[]? replacement)) {
                int length = checked((int)BitConverter.ToUInt32(bytes, chunk.Offset));
                writer.Write(replacement);
                writer.Write(bytes, chunk.Offset + length, boundary - chunk.Offset - length);
            } else writer.Write(bytes, chunk.Offset, boundary - chunk.Offset);
        }
        uint newTrailer = checked((uint)output.Position); writer.Write(tail);
        byte[] result = output.ToArray(); Write(result, 0x1A, newTrailer);
        return result;
    }

    private static byte[] Rewrite(byte[] publication, Action<Dictionary<string, byte[]>> mutation) {
        Assert.True(OfficeCompoundFileReader.TryRead(publication, out OfficeCompoundFile? compound, out string? error), error);
        var streams = new Dictionary<string, byte[]> {
            ["Contents"] = (byte[])compound!.Streams["Contents"].Clone(),
            ["Quill/QuillSub/CONTENTS"] = (byte[])compound.Streams["Quill/QuillSub/CONTENTS"].Clone()
        };
        mutation(streams);
        return OfficeCompoundFileWriter.Rewrite(compound, streams);
    }

    private static int ChunkOffset(byte[] bytes, uint id) {
        int trailer = checked((int)BitConverter.ToUInt32(bytes, 0x1A));
        Block directory = Chunk(bytes, trailer).Single(field => field.Type == 0x90);
        Block slot = Children(bytes, directory)[checked((int)id)];
        return checked((int)BitConverter.ToUInt32(bytes, Children(bytes, slot).Single(field => field.Id == 4).Offset));
    }
    private static Block[] Chunk(byte[] bytes, int start) => Blocks(bytes, start + 4, start + checked((int)BitConverter.ToUInt32(bytes, start)));
    private static Block[] Children(byte[] bytes, Block block) => Blocks(bytes, block.Offset + 4, block.Offset + block.Length);
    private static Block[] Blocks(byte[] bytes, int start, int end) {
        var result = new List<Block>();
        for (int position = start; position < end;) {
            byte type = bytes[position + 1]; int offset = position + 2;
            int length = type switch {
                0x00 or 0x02 or 0x05 or 0x08 or 0x0A or 0x78 => 0,
                0x07 or 0x10 or 0x12 or 0x18 or 0x1A => 2,
                0x20 or 0x22 or 0x58 or 0x68 or 0x70 or 0xB8 => 4,
                0x28 => 8, 0x38 => 16, 0x48 => 24,
                _ => checked((int)BitConverter.ToUInt32(bytes, offset))
            };
            result.Add(new Block(bytes[position], type, offset, length)); position = offset + length;
        }
        return result.ToArray();
    }
    private static (int Descriptor, int Data) TableTextChunk(byte[] bytes) {
        var chunks = new List<(int, int)>();
        uint next = 0x18;
        while (next != uint.MaxValue) {
            int directory = checked((int)next), count = BitConverter.ToUInt16(bytes, directory + 2);
            for (int i = 0; i < count; i++) {
                int position = directory + 8 + i * 24;
                if (Encoding.ASCII.GetString(bytes, position + 2, 4) == "TCD ")
                    chunks.Add((position, checked((int)BitConverter.ToUInt32(bytes, position + 16))));
            }
            next = BitConverter.ToUInt32(bytes, directory + 4);
        }
        return Assert.Single(chunks);
    }
    private static void Write(byte[] bytes, int offset, uint value) {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8));
    }
    private static void WriteScalar(byte[] bytes, Block field, uint value) {
        Assert.True(field.Length is 2 or 4);
        for (int i = 0; i < field.Length; i++) bytes[field.Offset + i] = (byte)(value >> (i * 8));
    }
    private readonly record struct Block(byte Id, byte Type, int Offset, int Length);
}
