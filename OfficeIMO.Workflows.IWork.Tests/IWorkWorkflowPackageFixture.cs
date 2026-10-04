using System.IO.Compression;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    private static void WritePackage(string path, byte[] records) {
        // One literal-only Snappy block inside the bounded IWA frame.
        int encodedLength = records.Length - 1;
        byte[] literal = records.Length <= 60 ? [(byte)(encodedLength << 2)]
            : encodedLength <= byte.MaxValue ? [(byte)(60 << 2), (byte)encodedLength]
            : [(byte)(61 << 2), (byte)encodedLength, (byte)(encodedLength >> 8)];
        byte[] block = Join(U((ulong)records.Length), literal, records);
        using var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        using Stream entry = zip.CreateEntry("Index/Document.iwa").Open();
        entry.Write(Join([0, (byte)block.Length, (byte)(block.Length >> 8), (byte)(block.Length >> 16)], block));
    }

    private static byte[] Join(params byte[][] values) => values.SelectMany(value => value).ToArray();
    private static byte[] U(ulong value) {
        var bytes = new List<byte>();
        do { byte next = (byte)(value & 127); value >>= 7; bytes.Add(value == 0 ? next : (byte)(next | 128)); } while (value != 0);
        return bytes.ToArray();
    }
    private static byte[] V(int field, ulong value) => Join(U((ulong)(field << 3)), U(value));
    private static byte[] B(int field, byte[] value) => Join(U((ulong)((field << 3) | 2)), U((ulong)value.Length), value);
    private static byte[] R(int field, ulong target) => B(field, V(1, target));
    private static byte[] A(ulong id, ulong type, byte[] value) {
        byte[] info = Join(V(1, id), B(2, Join(V(1, type), V(3, (ulong)value.Length))));
        return Join(U((ulong)info.Length), info, value);
    }
}
