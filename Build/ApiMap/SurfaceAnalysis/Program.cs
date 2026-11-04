using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Build.ApiMap;

if (args.Length < 2) {
    Console.Error.WriteLine("Usage: SurfaceAnalysis <repository-root> <output.json> [package ...]");
    return 2;
}
string root = Path.GetFullPath(args[0]);
string output = Path.GetFullPath(args[1]);
string[] packages = args.Skip(2).ToArray();
if (packages.Length == 0) packages = ["OfficeIMO.Core", "OfficeIMO.Word", "OfficeIMO.Excel", "OfficeIMO.PowerPoint", "OfficeIMO.Pdf", "OfficeIMO.Markdown", "OfficeIMO"];
var start = new ProcessStartInfo("git") { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true };
foreach (string argument in new[] { "ls-files", "--cached", "--others", "--exclude-standard" }) start.ArgumentList.Add(argument);
using var process = Process.Start(start) ?? throw new InvalidOperationException("Could not enumerate repository inputs.");
string input = process.StandardOutput.ReadToEnd();
string error = process.StandardError.ReadToEnd();
process.WaitForExit();
if (process.ExitCode != 0) throw new InvalidOperationException(error);
string[] files = input.Split('\n', StringSplitOptions.RemoveEmptyEntries).Distinct(StringComparer.Ordinal).ToArray();
var reports = SurfaceInventory.Create(root, files, packages);
Directory.CreateDirectory(Path.GetDirectoryName(output)!);
File.WriteAllText(output, JsonSerializer.Serialize(new { schemaVersion = 1, analysis = "Compiled public declarations, documentation references and cross-assembly model signature closure", packages = reports }, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase }) + "\n");
foreach (var report in reports) Console.WriteLine($"{report.Package}: {report.PublicTypes} public types; {report.Candidates.Count} candidates; {report.EvidenceFiles} documentation/example files");
return 0;
