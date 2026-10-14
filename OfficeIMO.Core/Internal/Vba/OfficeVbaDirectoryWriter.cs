using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Writes VBA directory records while retaining unedited and unknown reference/module metadata.</summary>
internal static class OfficeVbaDirectoryWriter {
    internal static byte[] CreateProject(string name, int codePage) {
        using var directory = new MemoryStream();
        Record(directory, 0x0001, BitConverter.GetBytes(1U));
        Record(directory, 0x0002, BitConverter.GetBytes(0x409U));
        Record(directory, 0x0014, BitConverter.GetBytes(0x409U));
        Record(directory, 0x0003, BitConverter.GetBytes((ushort)codePage));
        Record(directory, 0x0004, OfficeVbaText.Encode(name, codePage));
        foreach (ushort id in new ushort[] { 0x0005, 0x0040, 0x0006, 0x003d }) Record(directory, id, Array.Empty<byte>());
        Record(directory, 0x0007, BitConverter.GetBytes(0U));
        Record(directory, 0x0008, BitConverter.GetBytes(0U));
        // PROJECTVERSION's size covers the four-byte major version; its minor follows outside that size.
        Record(directory, 0x0009, BitConverter.GetBytes(1U));
        directory.WriteByte(0); directory.WriteByte(0);
        Record(directory, 0x000c, Array.Empty<byte>());
        Record(directory, 0x003c, Array.Empty<byte>());
        Record(directory, 0x000f, BitConverter.GetBytes((ushort)0));
        Record(directory, 0x0013, BitConverter.GetBytes((ushort)0));
        Record(directory, 0x0010, Array.Empty<byte>());
        var streams = new List<OfficeCompoundStream> {
            new("PROJECT", OfficeVbaText.Encode(OfficeVbaProjectText.Create(name), codePage)),
            new("PROJECTwm", new byte[] { 0, 0 }),
            new("VBA/dir", OfficeVbaCompression.Compress(directory.ToArray())),
            new("VBA/_VBA_PROJECT", new byte[] { 0xcc, 0x61, 0xff, 0xff, 0, 1, 0 })
        };
        return OfficeCompoundFileWriter.Write(streams);
    }

    internal static OfficeVbaDirectoryCodec.ModuleModel NewModule(string name, OfficeVbaModuleKind kind, int codePage) {
        using var output = new MemoryStream();
        byte[] ansi = OfficeVbaText.Encode(name, codePage);
        byte[] unicode = Encoding.Unicode.GetBytes(name);
        Record(output, 0x0019, ansi);
        Record(output, 0x0047, unicode);
        Record(output, 0x001a, ansi);
        Record(output, 0x0032, unicode);
        Record(output, 0x001c, Array.Empty<byte>());
        Record(output, 0x0048, Array.Empty<byte>());
        Record(output, 0x0031, BitConverter.GetBytes(0U));
        Record(output, 0x001e, BitConverter.GetBytes(0U));
        Record(output, 0x002c, BitConverter.GetBytes((ushort)0xffff));
        ushort type = kind == OfficeVbaModuleKind.Standard ? (ushort)0x0021 : (ushort)0x0022;
        Record(output, type, Array.Empty<byte>());
        Record(output, 0x002b, Array.Empty<byte>());
        return new OfficeVbaDirectoryCodec.ModuleModel {
            Serialized = output.ToArray(), AnsiName = ansi, UnicodeName = unicode, StreamName = name,
            TextOffset = 0, TypeId = type
        };
    }

    internal static byte[] SerializeModule(OfficeVbaModule module, int codePage, bool sourceChanged) {
        byte[] original = module.Directory.Serialized;
        if (!sourceChanged && module.Name == module.OriginalName) return original;
        using var input = new MemoryStream(original, writable: false);
        using var reader = new BinaryReader(input);
        using var output = new MemoryStream();
        while (input.Position < input.Length) {
            ushort id = reader.ReadUInt16();
            uint size = reader.ReadUInt32();
            if (size > int.MaxValue || size > input.Length - input.Position) throw new InvalidDataException("The module directory record is truncated.");
            byte[] bytes = reader.ReadBytes((int)size);
            if (module.Name != module.OriginalName) {
                if (id == 0x0019 || id == 0x001a) bytes = OfficeVbaText.Encode(module.Name, codePage);
                else if (id == 0x0047 || id == 0x0032) bytes = Encoding.Unicode.GetBytes(module.Name);
            }
            if (sourceChanged && id == 0x0031) bytes = BitConverter.GetBytes(0U);
            Record(output, id, bytes);
        }
        return output.ToArray();
    }

    internal static byte[] SerializeDirectory(OfficeVbaDirectoryCodec.DirectoryModel directory,
        IReadOnlyList<OfficeVbaModule> modules, IReadOnlyList<OfficeVbaReference> references, int codePage) {
        using var output = new MemoryStream();
        output.Write(directory.SerializedPrefix, 0, directory.SerializedPrefix.Length);
        foreach (OfficeVbaReference reference in references) output.Write(reference.Serialized, 0, reference.Serialized.Length);
        Record(output, 0x000f, BitConverter.GetBytes(checked((ushort)modules.Count)));
        output.Write(directory.SerializedCookie, 0, directory.SerializedCookie.Length);
        foreach (OfficeVbaModule module in modules) {
            byte[] record = SerializeModule(module, codePage, module.IsNew || module.Source != module.OriginalSource || module.Name != module.OriginalName);
            output.Write(record, 0, record.Length);
        }
        output.Write(directory.TerminatorRecord, 0, directory.TerminatorRecord.Length);
        return output.ToArray();
    }

    internal static byte[] RegisteredReference(string name, string libraryId, int codePage) {
        using var output = new MemoryStream();
        Record(output, 0x0016, OfficeVbaText.Encode(name, codePage));
        Record(output, 0x003e, Encoding.Unicode.GetBytes(name));
        byte[] libid = OfficeVbaText.Encode(libraryId, codePage);
        using var body = new MemoryStream();
        using (var writer = new BinaryWriter(body, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(libid.Length); writer.Write(libid); writer.Write(0U); writer.Write((ushort)0);
        }
        Record(output, 0x000d, body.ToArray());
        return output.ToArray();
    }

    internal static OfficeVbaReference ReadReference(byte[] serialized, int codePage) {
        string name = string.Empty;
        string? libid = null;
        using var input = new MemoryStream(serialized, writable: false);
        using var reader = new BinaryReader(input);
        if (input.Length >= 6 && reader.ReadUInt16() == 0x0016) {
            int count = checked((int)reader.ReadUInt32());
            name = OfficeVbaText.Decode(reader.ReadBytes(count), codePage);
            if (reader.ReadUInt16() != 0x003e) throw new InvalidDataException("The reference Unicode name record is missing.");
            int unicodeCount = checked((int)reader.ReadUInt32());
            byte[] unicode = reader.ReadBytes(unicodeCount);
            if (unicode.Length > 0) name = new UnicodeEncoding(false, false, true).GetString(unicode);
        } else input.Position = 0;
        if (input.Length - input.Position >= 10 && reader.ReadUInt16() == 0x000d) {
            reader.ReadUInt32();
            int count = checked((int)reader.ReadUInt32());
            libid = OfficeVbaText.Decode(reader.ReadBytes(count), codePage);
        }
        return new OfficeVbaReference(name, libid, (byte[])serialized.Clone());
    }

    internal static byte[] ProjectNames(IReadOnlyList<OfficeVbaModule> modules, int codePage) {
        using var output = new MemoryStream();
        foreach (OfficeVbaModule module in modules) {
            byte[] ansi = OfficeVbaText.Encode(module.Name, codePage);
            byte[] unicode = Encoding.Unicode.GetBytes(module.Name);
            output.Write(ansi, 0, ansi.Length); output.WriteByte(0);
            output.Write(unicode, 0, unicode.Length); output.WriteByte(0); output.WriteByte(0);
        }
        output.WriteByte(0); output.WriteByte(0);
        return output.ToArray();
    }

    private static void Record(Stream output, ushort id, byte[] bytes) {
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        writer.Write(id); writer.Write(bytes.Length); writer.Write(bytes);
    }
}
