using MPXJ.Net;
using System.Text.Json;

// An opt-in independent file reader/writer oracle. This project is intentionally
// outside the OfficeIMO solution and has no runtime or package consumers.
if (args.Length == 3 && args[0] == "batch-mspdi") {
    string directory = Path.GetFullPath(args[1]);
    var inputs = Directory.GetFiles(directory, args[2]).Where(p => new[] { ".mpp", ".mpt", ".mpx" }.Contains(Path.GetExtension(p).ToLowerInvariant())).OrderBy(p => p).ToArray();
    string reportPath = Path.Combine(directory, "independent-batch.json");
    if (File.Exists(reportPath) || inputs.Any(p => File.Exists(Path.ChangeExtension(p, ".xml")))) throw new IOException("Independent batch outputs must not already exist.");
    var results = new List<object>(); int failures = 0;
    foreach (string file in inputs) {
        try {
            var project = new UniversalProjectReader().Read(file);
            new UniversalProjectWriter(FileFormat.MSPDI).Write(project, Path.ChangeExtension(file, ".xml"));
            results.Add(new { file = Path.GetFileName(file), passed = true });
        } catch (Exception error) { failures++; results.Add(new { file = Path.GetFileName(file), passed = false, error = error.Message }); }
    }
    File.WriteAllText(reportPath, JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine("Independent MSPDI exports=" + inputs.Length + " failures=" + failures);
    return failures == 0 ? 0 : 1;
}
if (args.Length != 3) {
    Console.Error.WriteLine("Usage: <input-file> <new-output-file> <MSPDI|MPX|JSON>");
    return 2;
}
string input = Path.GetFullPath(args[0]), output = Path.GetFullPath(args[1]);
if (File.Exists(output)) throw new IOException("Choose a new output file.");
var format = args[2] switch {
    "MSPDI" => FileFormat.MSPDI,
    "MPX" => FileFormat.MPX,
    "JSON" => FileFormat.JSON,
    _ => throw new ArgumentException("Choose MSPDI, MPX, or JSON.")
};
try {
    var project = new UniversalProjectReader().Read(input);
    new UniversalProjectWriter(format).Write(project, output);
    Console.WriteLine("Independent reader accepted " + Path.GetFileName(input) + "; wrote " + Path.GetFileName(output));
    return 0;
} catch (Exception error) {
    if (error is java.lang.Throwable throwable) throwable.printStackTrace();
    else Console.Error.WriteLine(error);
    return 1;
}
