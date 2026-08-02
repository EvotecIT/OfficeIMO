using System.Text;
using System.Text.Json;

string root = FindRoot(Directory.GetCurrentDirectory());
string source = GetOption(args, "--source") ?? Path.Combine(root, "Build", "AcceptanceManifest", "acceptance-manifest.json");
string output = GetOption(args, "--output") ?? Path.Combine(root, "Docs", "Compatibility", "generated", "email-cloud-acceptance.md");
bool verify = args.Contains("--verify", StringComparer.OrdinalIgnoreCase);
AcceptanceManifest manifest = JsonSerializer.Deserialize<AcceptanceManifest>(File.ReadAllText(source), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
    ?? throw new InvalidDataException("Acceptance manifest could not be read.");
if (manifest.SchemaVersion != 1 || manifest.Rows.Count == 0) throw new InvalidDataException("Acceptance manifest schema or rows are invalid.");
string? duplicate = manifest.Rows.GroupBy(row => row.Id, StringComparer.Ordinal).Where(group => group.Count() > 1).Select(group => group.Key).FirstOrDefault();
if (duplicate != null) throw new InvalidDataException("Duplicate acceptance id: " + duplicate);
foreach (AcceptanceRow row in manifest.Rows) {
    string[] parts = row.Evidence.Split(new[] { "::" }, 2, StringSplitOptions.None);
    string path = Path.Combine(root, parts[0].Replace('/', Path.DirectorySeparatorChar));
    if (!File.Exists(path)) throw new InvalidDataException($"Evidence file does not exist for {row.Id}: {parts[0]}");
    if (parts.Length == 2 && File.ReadAllText(path).IndexOf(parts[1], StringComparison.Ordinal) < 0) throw new InvalidDataException($"Evidence symbol does not exist for {row.Id}: {parts[1]}");
}
string markdown = Render(manifest);
if (verify) {
    if (!File.Exists(output) || Normalize(File.ReadAllText(output)) != Normalize(markdown)) { Console.Error.WriteLine("Generated email/cloud acceptance evidence is missing or stale."); Environment.ExitCode = 1; }
    else Console.WriteLine("Verified generated email/cloud acceptance evidence.");
} else { Directory.CreateDirectory(Path.GetDirectoryName(output)!); File.WriteAllText(output, markdown, new UTF8Encoding(false)); Console.WriteLine("Generated " + output); }

static string Render(AcceptanceManifest manifest) {
    var text = new StringBuilder();
    text.AppendLine("# Email, stores, and cloud acceptance evidence").AppendLine();
    text.AppendLine("> Generated from `Build/AcceptanceManifest/acceptance-manifest.json`. Edit the manifest and regenerate; do not hand-edit this file.").AppendLine();
    text.AppendLine("| Area | Target | Operation | Status | Evidence | Current boundary |");
    text.AppendLine("| --- | --- | --- | --- | --- | --- |");
    foreach (AcceptanceRow row in manifest.Rows) text.Append("| ").Append(E(row.Area)).Append(" | ").Append(E(row.Target)).Append(" | ").Append(E(row.Operation)).Append(" | ").Append(E(row.Status)).Append(" | `").Append(E(row.Evidence)).Append("` | ").Append(E(row.Notes)).AppendLine(" |");
    return Normalize(text.ToString()).TrimEnd('\n') + "\n";
}
static string E(string value) => value.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
static string Normalize(string value) => value.Replace("\r\n", "\n").Replace("\r", "\n");
static string? GetOption(string[] args, string name) { for (int i = 0; i < args.Length; i++) if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return i + 1 < args.Length ? args[i + 1] : throw new ArgumentException(name + " requires a value."); return null; }
static string FindRoot(string start) { DirectoryInfo? directory = new DirectoryInfo(start); while (directory != null && !File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) directory = directory.Parent; return directory?.FullName ?? throw new DirectoryNotFoundException("OfficeIMO repository root not found."); }
internal sealed class AcceptanceManifest { public int SchemaVersion { get; set; } public List<AcceptanceRow> Rows { get; set; } = new(); }
internal sealed class AcceptanceRow { public string Id { get; set; } = ""; public string Area { get; set; } = ""; public string Target { get; set; } = ""; public string Operation { get; set; } = ""; public string Status { get; set; } = ""; public string Evidence { get; set; } = ""; public string Notes { get; set; } = ""; }
