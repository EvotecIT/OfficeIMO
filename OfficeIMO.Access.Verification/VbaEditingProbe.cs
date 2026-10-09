using OfficeIMO;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Access.Verification {
    /// <summary>Executable public native VBA workflow over independently produced synthetic applications.</summary>
    internal static class VbaEditingProbe {
        internal static void Generate(string fixtures, string outputDirectory) {
            string root = Path.GetFullPath(outputDirectory);
            Directory.CreateDirectory(root);
            var observations = new List<object>();
            foreach (string file in new[] { "Application/objects-jet4.mdb", "Application/objects-ace12.accdb", "Designer/designer-jet4.mdb", "Designer/designer-ace12.accdb" }) {
                string input = Path.Combine(fixtures, file);
                byte[] original = File.ReadAllBytes(input);
                using AccessDocument database = AccessDocument.Load(input);
                OfficeVbaProject project = database.GetVbaProject();
                string name = project.Modules.Single().Name;
                string source = project.GetModule(name).Source.Replace("Value = 42", "Value = 43");
                if (source == project.GetModule(name).Source) throw new InvalidDataException("The native fixture source marker is absent.");
                // The native oracle runs on Windows ANSI 1252. These characters have
                // identical bytes in both 1250 and 1252; portable tests retain Polish text.
                source = source.Replace("'Zażółć gęślą jaźń", "'Native source: éóö€");
                project.SetModuleSource(name, source);
                string growth = string.Join("\r\n", Enumerable.Range(0, 1800).Select(i => "' " + Convert.ToHexString(SHA256.HashData(BitConverter.GetBytes(i)))));
                project.AddModule("AddedModule", "Option Explicit\r\nPublic Function AddedValue() As Long\r\n AddedValue = 44\r\nEnd Function\r\n" + growth);
                project.AddModule("AddedClass", "Option Explicit\r\nPublic Function ClassValue() As Long\r\n ClassValue = 45\r\nEnd Function\r\n", OfficeVbaModuleKind.Class);
                database.SetVbaProject(project);
                string output = Path.Combine(root, Path.GetFileName(file));
                database.Save(output);
                File.WriteAllText(output + ".expected-module.txt", project.GetModule(name).Source);
                File.WriteAllText(output + ".expected-added.txt", project.GetModule("AddedModule").Source);
                using AccessDocument saved = AccessDocument.Load(output);
                foreach (OfficeVbaModule module in project.Modules) {
                    if (saved.GetVbaProject().GetModule(module.Name).Source.TrimEnd('\r', '\n') != module.Source.TrimEnd('\r', '\n'))
                        throw new InvalidDataException("Persisted VBA source differs: " + module.Name);
                }
                foreach (AccessStorageStream stream in database.ApplicationStreams.Where(x => !x.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase) && !x.Path.StartsWith("Modules/", StringComparison.OrdinalIgnoreCase))) {
                    if (!stream.Payload.GetBytes().SequenceEqual(saved.ApplicationStreams.Single(x => x.Path == stream.Path).Payload.GetBytes()))
                        throw new InvalidDataException("Unrelated native stream differs: " + stream.Path);
                }
                if (!original.SequenceEqual(File.ReadAllBytes(input))) throw new InvalidDataException("The input fixture changed.");
                observations.Add(new { file, output, module = name, profile = saved.Profile.ToString(), sourceExact = true, unrelatedStreamsExact = true,
                    inputUnchanged = true, outputBytes = new FileInfo(output).Length, sha256 = saved.Inspection!.Sha256 });
            }
            File.WriteAllText(Path.Combine(root, "vba-editing.json"), JsonSerializer.Serialize(observations, new JsonSerializerOptions { WriteIndented = true }));
        }
    }
}
