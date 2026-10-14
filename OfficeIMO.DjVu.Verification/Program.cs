using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.DjVu;
using OfficeIMO.DjVu.Verification;

if (args.Length != 3) {
    Console.Error.WriteLine("Usage: OfficeIMO.DjVu.Verification <input.djvu> <explicit-ddjvu-executable> <report.json>");
    return 2;
}
string input = Path.GetFullPath(args[0]), oracle = Path.GetFullPath(args[1]), report = Path.GetFullPath(args[2]);
if (report == input || report == oracle) throw new ArgumentException("The report must have a separate output path.");
var document = DjVuDocument.Load(input);
var evidence = new RenderingEvidence { InputName = Path.GetFileName(input), SourceSha256 = document.SourceSha256,
    SourceBytes = document.SourceLengthBytes, OracleVersion = await DjVuReferenceRenderer.GetVersionAsync(oracle) };
for (int i = 0; i < document.Pages.Count; i++) {
    try {
        var page = document.Pages[i];
        var image = page.Render().Image;
        var comparison = await DjVuReferenceRenderer.CompareAsync(oracle, input, page.Number, image);
        comparison.PageNumber = page.Number;
        comparison.Dpi = page.Dpi;
        comparison.StoredTextStatus = page.GetText().Status.ToString();
        evidence.Pages.Add(comparison);
        if (!comparison.Passed) Console.WriteLine($"Page {page.Number}: max {comparison.MaximumDifference.Max()}, mean {comparison.MeanDifference.Max():F6}");
    } catch (Exception error) {
        evidence.Pages.Add(new PageRenderingEvidence { PageNumber = i + 1, Error = error.GetType().Name + ": " + error.Message });
    }
    if ((i + 1) % 8 == 0 || i + 1 == document.Pages.Count) {
        File.WriteAllText(report + ".tmp", JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
        File.Move(report + ".tmp", report, true);
        Console.WriteLine($"{i + 1}/{document.Pages.Count} pages; {evidence.Pages.Count(p => !p.Passed)} failures");
    }
}
return evidence.Pages.All(p => p.Passed) ? 0 : 1;
