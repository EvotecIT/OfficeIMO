using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using HtmlForgeX;
using SyntheticReportProducer;

var options = ReadOptions(args);
var contractPath = RequiredPath(options, "contract");
var inputsPath = RequiredPath(options, "inputs");
var outputPath = RequiredPath(options, "output");
var contractBytes = File.ReadAllBytes(contractPath);
var contract = JsonSerializer.Deserialize<CorpusSpec>(contractBytes,
    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
    ?? throw new InvalidOperationException("The predeclared fixture contract is empty.");
if (contract.Classification != "synthetic-independent-producer") {
    throw new InvalidOperationException("Only the declared synthetic corpus is supported.");
}
var sourceBindingBytes = File.ReadAllBytes(Path.Combine(
    Path.GetDirectoryName(contractPath)!, contract.SourceBinding));

var css = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "report.css"));
foreach (var font in contract.Fonts) {
    var bytes = File.ReadAllBytes(Path.Combine(inputsPath, font.FileName));
    var hash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    if (!StringComparer.Ordinal.Equals(hash, font.Sha256)) {
        throw new InvalidOperationException($"Font content differs from the declared input: {font.FileName}");
    }
    css = css.Replace($"@@{font.Role.ToUpperInvariant()}_FONT_BASE64@@", Convert.ToBase64String(bytes), StringComparison.Ordinal);
}

Directory.CreateDirectory(outputPath);
var files = new List<object>();
foreach (var report in contract.Reports) {
    foreach (var variant in contract.Variants) {
        var finalCss = css
            .Replace("@@PAGE_ORIENTATION@@", variant.Orientation, StringComparison.Ordinal)
            .Replace("@@SCREEN_WIDTH@@", variant.ScreenWidth, StringComparison.Ordinal);
        var fileName = $"{report.Key}-{variant.Key}.html";
        var html = ReportComposer.Compose(report, variant, finalCss);
        File.WriteAllText(Path.Combine(outputPath, fileName), html, new UTF8Encoding(false));
        files.Add(new {
            fileName,
            report = report.Key,
            variant = variant.Key,
            bytes = Encoding.UTF8.GetByteCount(html),
            sha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(html))).ToLowerInvariant()
        });
        Console.WriteLine(fileName);
    }
}

// This receipt binds the generated inputs. Expected content remains in the input contract;
// the generator never derives, updates, or overwrites that inventory from rendered output.
var receipt = new {
    classification = contract.Classification,
    contractSha256 = Convert.ToHexString(SHA256.HashData(contractBytes)).ToLowerInvariant(),
    producerAssembly = typeof(Table).Assembly.GetName().FullName,
    sourceBindingFile = contract.SourceBinding,
    sourceBindingBase = "contract-directory",
    sourceBindingSha256 = Convert.ToHexString(SHA256.HashData(sourceBindingBytes)).ToLowerInvariant(),
    emission = "Typed HtmlForgeX component fragments in a local HTML shell; no Document library registration or scripts.",
    files
};
File.WriteAllText(Path.Combine(outputPath, "producer-receipt.json"),
    JsonSerializer.Serialize(receipt, new JsonSerializerOptions { WriteIndented = true }) + "\n",
    new UTF8Encoding(false));

static Dictionary<string, string> ReadOptions(string[] arguments) {
    if (arguments.Length != 6) {
        throw new ArgumentException("Usage: --contract <json> --inputs <font-directory> --output <report-directory>");
    }
    var values = new Dictionary<string, string>(StringComparer.Ordinal);
    for (var i = 0; i < arguments.Length; i += 2) {
        if (!arguments[i].StartsWith("--", StringComparison.Ordinal)) {
            throw new ArgumentException("Expected a named option.");
        }
        values.Add(arguments[i][2..], arguments[i + 1]);
    }
    return values;
}

static string RequiredPath(IReadOnlyDictionary<string, string> values, string key) {
    if (!values.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value)) {
        throw new ArgumentException($"Missing --{key}.");
    }
    return Path.GetFullPath(value);
}
