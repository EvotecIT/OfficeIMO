using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeIMO;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Rtf.Pdf;

// Converts one PDF to each requested editable or semantic target with the public OfficeIMO APIs.
// The Python driver owns parallelism, timeouts and scoring; this process only converts and records outcomes.
if (args.Length < 4 || args[0] != "convert") {
    Console.Error.WriteLine("usage: convert <input.pdf> <output-directory> <targets: docx,xlsx,pptx,rtf,odt,ods,odp,html,md>");
    return 2;
}

string input = Path.GetFullPath(args[1]);
string outputDirectory = Path.GetFullPath(args[2]);
string[] targets = args[3].Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
Directory.CreateDirectory(outputDirectory);
string stem = Path.GetFileNameWithoutExtension(input);
var results = new List<Dictionary<string, object?>>();

foreach (string target in targets) {
    string output = Path.Combine(outputDirectory, stem + "." + target);
    var entry = new Dictionary<string, object?> { ["target"] = target, ["output"] = output };
    var timer = Stopwatch.StartNew();
    try {
        // Each target loads the PDF afresh so one converter cannot affect another through shared state.
        using PdfDocument document = PdfDocument.Load(input);
        OfficeConversionReport? report = target switch {
            "docx" => document.SaveAsWord(output),
            "xlsx" => document.SaveAsExcel(output),
            "pptx" => document.SaveAsPowerPoint(output),
            "rtf" => document.SaveAsRtf(output),
            "odt" => document.SaveAsOdt(output),
            "ods" => document.SaveAsOds(output),
            "odp" => document.SaveAsOdp(output),
            "html" => document.SaveAsHtml(output),
            "md" => WriteMarkdown(document, output),
            _ => throw new ArgumentException($"Unknown target '{target}'.")
        };
        entry["status"] = "ok";
        if (report != null) {
            entry["diagnostics"] = report.Diagnostics.Count;
            entry["lossDiagnostics"] = report.Diagnostics.Count(static d => d.LossKind != OfficeConversionLossKind.None);
            entry["diagnosticCodes"] = report.Diagnostics.Select(static d => d.Code).Distinct().Take(20).ToArray();
        }
    } catch (Exception exception) {
        entry["status"] = "error";
        entry["error"] = exception.GetType().Name + ": " + Truncate(exception.Message, 400);
        try { if (File.Exists(output)) File.Delete(output); } catch (IOException) { }
    }
    entry["seconds"] = Math.Round(timer.Elapsed.TotalSeconds, 3);
    results.Add(entry);
}

File.WriteAllText(Path.Combine(outputDirectory, stem + ".officeimo.json"),
    JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
return 0;

static OfficeConversionReport? WriteMarkdown(PdfDocument document, string output) {
    File.WriteAllText(output, document.ExtractMarkdown(), new UTF8Encoding(false));
    return null;
}

static string Truncate(string value, int length) => value.Length <= length ? value : value.Substring(0, length);
