using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security.Tests;

internal static class OfficeVbaMalformedTextFixtures {
    internal static byte[] Create(string field) {
        var project = OfficeVbaProject.Create("Inventory", 65001);
        project.AddModule("Helpers", "'source");
        project.AddRegisteredReference("Library", new Guid("11111111-2222-3333-4444-555555555555"));
        Assert.True(OfficeCompoundFileReader.TryRead(project.Write().GetBytes(), out OfficeCompoundFile? compound, out _));
        var replacements = new Dictionary<string, byte[]>();
        if (field == "project") replacements["PROJECT"] = new byte[] { 0xff };
        else if (field == "source") replacements["VBA/Helpers"] = OfficeVbaCompression.Compress(new byte[] { 0xff });
        else {
            Assert.True(OfficeVbaCompression.TryDecompress(compound!.Streams["VBA/dir"], 20000, out byte[] directory, out _));
            switch (field) {
                case "module": Corrupt(directory, Record(0x0047, Encoding.Unicode.GetBytes("Helpers")), 6, 0, 0xd8); break;
                case "projectName": Corrupt(directory, Record(0x0004, Encoding.UTF8.GetBytes("Inventory")), 6, 0xff); break;
                case "referenceName": Corrupt(directory, Record(0x003e, Encoding.Unicode.GetBytes("Library")), 6, 0, 0xd8); break;
                case "referenceId": Corrupt(directory, Encoding.UTF8.GetBytes(project.References.Single().LibraryId!), 0, 0xff); break;
                case "unsupported": Corrupt(directory, Record(0x0003, BitConverter.GetBytes((ushort)65001)), 6, 0xff, 0xff); break;
                default: throw new ArgumentException("Unknown malformed field.", nameof(field));
            }
            replacements["VBA/dir"] = OfficeVbaCompression.Compress(directory);
        }
        return OfficeCompoundFileWriter.Rewrite(compound!, replacements);
    }

    private static byte[] Record(ushort id, byte[] value) => BitConverter.GetBytes(id)
        .Concat(BitConverter.GetBytes(value.Length)).Concat(value).ToArray();

    private static void Corrupt(byte[] bytes, byte[] marker, int offset, params byte[] replacement) {
        int position = Enumerable.Range(0, bytes.Length - marker.Length + 1)
            .Single(index => bytes.Skip(index).Take(marker.Length).SequenceEqual(marker));
        Array.Copy(replacement, 0, bytes, position + offset, replacement.Length);
    }
}
