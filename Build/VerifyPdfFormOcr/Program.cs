using System.Text.Json;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

// Opt-in consumer proof. A JSON decision file supplies explicitly reviewed values;
// neither recognition nor confidence creates acceptance automatically.
if (args.Length != 4) throw new ArgumentException("Usage: VerifyPdfFormOcr <source.pdf> <output-directory> <accepted-values.json> <tesseract-executable>");
string sourcePath = Path.GetFullPath(args[0]), folder = Path.GetFullPath(args[1]);
if (Directory.Exists(folder)) throw new ArgumentException("Use a new task-owned output directory.");
Directory.CreateDirectory(folder);
string output = Path.Combine(folder, "reviewed-form.pdf");
byte[] original = File.ReadAllBytes(sourcePath);
var source = PdfDocument.Load(original);
var engine = TesseractOcrEngine.CreateDefault(new() { ExecutablePath = args[3] });
var review = await new OfficeWorkflowRunner().PreparePdfFormOcrAsync(source, engine, new() {
    Language = "eng", Dpi = 144, MaxPages = 5, MaxPixelsPerPage = 10_000_000,
    MinimumConfidence = 0.75, MaxOcrTextCharactersPerPage = 100_000
});
var values = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(args[2]))
    ?? throw new ArgumentException("Provide a JSON object with existing field names and explicitly reviewed values.");
var accepted = values.ToDictionary(pair => review.Proposals.Single(proposal => proposal.Field.Name == pair.Key),
    pair => PdfFormFieldValue.From(pair.Value));
var result = review.Apply(source, accepted);
result.Save(output).RequireSuccess();
var readback = PdfDocument.Load(output).Inspect();
var report = new {
    review.SourceFingerprint, Provider = engine.Id,
    RecognitionAccuracy = "Bounded fixture observation only; no general accuracy claim.",
    Proposals = review.Proposals.Select(proposal => new {
        proposal.Field.Name, proposal.SuggestedValue, proposal.Confidence,
        proposal.HasLowConfidence, proposal.IsAmbiguous, proposal.RejectionReason,
        Accepted = accepted.ContainsKey(proposal),
        SavedValues = readback.FormFieldsByName[proposal.Field.Name!].Values,
        Evidence = proposal.Evidence.Select(item => new { item.PageNumber, item.Word.Word.Text,
            item.Word.Word.Confidence, item.Word.Disposition, item.WidgetBounds })
    }),
    SourceUnchanged = original.SequenceEqual(File.ReadAllBytes(sourcePath))
};
File.WriteAllText(Path.Combine(folder, "review.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
if (!original.SequenceEqual(File.ReadAllBytes(sourcePath))) throw new InvalidOperationException("Source bytes changed.");
Console.WriteLine($"Reviewed {review.Proposals.Count} proposals; applied {accepted.Count} explicit values to {output}.");
