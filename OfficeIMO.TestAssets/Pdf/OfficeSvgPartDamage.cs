using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Tests;

/// <summary>Damages the SVG parts of a saved Office package while keeping each picture's PNG fallback intact.</summary>
internal static class OfficeSvgPartDamage {
    /// <summary>Truncates every <c>.svg</c> part so it no longer parses, then rewinds the package stream.</summary>
    internal static void Truncate(MemoryStream package) {
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".svg", StringComparison.OrdinalIgnoreCase)).ToList()) {
                string name = entry.FullName;
                entry.Delete();
                using Stream content = archive.CreateEntry(name).Open();
                byte[] damaged = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><rect width='40'");
                content.Write(damaged, 0, damaged.Length);
            }
        }
        package.Position = 0;
    }
}
