using System.IO.Compression;

namespace OfficeIMO.Tests;

/// <summary>Builds packages dominated by one large part, the way a large picture dominates a real file.</summary>
internal static class LargePartPackages {
    /// <summary>
    /// Adds an unreferenced, incompressible part of <paramref name="partBytes"/> bytes stored without compression, so
    /// the file is about that size and loading cost can be compared with it.
    /// </summary>
    internal static byte[] WithLargePart(byte[] package, string folder, int partBytes = 8 * 1024 * 1024) {
        using var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry types = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidDataException("The package has no content types part.");
            string xml;
            using (var reader = new StreamReader(types.Open())) xml = reader.ReadToEnd();
            if (xml.IndexOf("Extension=\"bin\"", StringComparison.OrdinalIgnoreCase) < 0) {
                int first = xml.IndexOf("<Default ", StringComparison.Ordinal);
                xml = xml.Insert(first, "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>");
                types.Delete();
                using var writer = new StreamWriter(archive.CreateEntry("[Content_Types].xml").Open());
                writer.Write(xml);
            }

            var part = new byte[partBytes];
            new Random(partBytes).NextBytes(part);
            using Stream stream = archive.CreateEntry(folder + "/media/large.bin", CompressionLevel.NoCompression).Open();
            stream.Write(part, 0, part.Length);
        }
        return output.ToArray();
    }
}
