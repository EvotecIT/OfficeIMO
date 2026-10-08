using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Access;

namespace OfficeIMO.Access.Verification;

internal static class PreservationProbe {
    internal static void Generate(string corpusDirectory, string outputDirectory) {
        string root = Path.GetFullPath(outputDirectory);
        if (Directory.Exists(root) || File.Exists(root)) throw new IOException("Preservation qualification requires a fresh output directory.");
        var records = new List<object>();
        foreach (string file in new[] { "Application/objects-jet4.mdb", "Application/objects-ace12.accdb", "Designer/designer-jet4.mdb", "Designer/designer-ace12.accdb" }) {
            string source = Path.Combine(corpusDirectory, file); string target = Path.Combine(root, Path.GetFileName(file));
            using var document = AccessDocument.Load(source);
            document.AssessSave(target).RequireNoLoss(); document.Save(target);
            byte[] original = File.ReadAllBytes(source), copy = File.ReadAllBytes(target);
            if (!original.SequenceEqual(copy)) throw new InvalidDataException("No-op native output differs from its source snapshot.");
            records.Add(new { file = Path.GetFileName(file), source = file, sha256 = Convert.ToHexString(SHA256.HashData(copy)).ToLowerInvariant(),
                applicationStreams = document.ApplicationStreams.Select(x => new { x.Path, length = x.Payload.Length, sha256 = Convert.ToHexString(SHA256.HashData(x.Payload.GetBytes())).ToLowerInvariant() }).ToArray(),
                vbaModules = document.VbaProject.ModuleNames, forms = document.Forms.Select(x => x.Name).ToArray(), reports = document.Reports.Select(x => x.Name).ToArray(), macros = document.Macros.Select(x => x.Name).ToArray(), dataMacros = document.DataMacros.Select(x => x.Event).ToArray() });
        }
        File.WriteAllText(Path.Combine(root, "preservation.json"), JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true }));
    }
}
