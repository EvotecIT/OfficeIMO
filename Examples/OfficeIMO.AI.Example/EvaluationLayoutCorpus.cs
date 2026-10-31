using System.Text;
using OfficeIMO.AI;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

internal static class EvaluationLayoutCorpus {
    public static IReadOnlyList<EvaluationCase> Create() {
        var font = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(16) };
        if (!font.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) throw new InvalidOperationException("Evaluation requires an embeddable font.");
        byte[] columns = PdfDocument.Create(builder => builder.Content(content => content.Canvas(canvas => {
            canvas.Text("Regional dispatches", global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(480), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(22));
            canvas.Text("WEST REGION", global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(35), global::OfficeIMO.OfficeLength.Points(18));
            canvas.Text("EAST REGION", global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(35), global::OfficeIMO.OfficeLength.Points(18));
            canvas.Text("Reference: W-731", global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(165), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(30));
            canvas.Text("Reference: E-294", global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(165), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(30));
            canvas.Text("Crates: 31", global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(205), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(30));
            canvas.Text("Crates: 98", global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(205), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(30));
            canvas.Text("Arrival: Tuesday", global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(245), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(30));
            canvas.Text("Arrival: Thursday", global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(245), global::OfficeIMO.OfficeLength.Points(220), global::OfficeIMO.OfficeLength.Points(30));
        })), font).ToBytes();
        var fields = new OfficeAiRequest { Operation = OfficeAiOperation.ExtractFields,
            Instruction = "Extract westReference, westCrates, eastReference and eastCrates, preserving the correct regional associations.",
            Fields = new[] { new OfficeAiFieldDefinition("westReference"), new OfficeAiFieldDefinition("westCrates", OfficeAiFieldType.Integer),
                new OfficeAiFieldDefinition("eastReference"), new OfficeAiFieldDefinition("eastCrates", OfficeAiFieldType.Integer) } };
        var columnGold = new EvaluationGold(Fields: new[] {
            new EvaluationFieldGold("westReference", OfficeAiFieldStatus.Present, "W-731"), new EvaluationFieldGold("westCrates", OfficeAiFieldStatus.Present, "31"),
            new EvaluationFieldGold("eastReference", OfficeAiFieldStatus.Present, "E-294"), new EvaluationFieldGold("eastCrates", OfficeAiFieldStatus.Present, "98") });
        byte[] diagram = PdfDocument.Create(builder => builder.Content(content => content.Canvas(canvas => {
            canvas.Text("Record processing", global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(480), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(22));
            for (int index = 0; index < 3; index++) {
                double x = 45 + 180 * index;
                canvas.Shape(OfficeShape.Rectangle(125, 60), global::OfficeIMO.OfficeLength.Points(x), global::OfficeIMO.OfficeLength.Points(150));
                canvas.Text(new[] { "Intake", "Review", "Archive" }[index], global::OfficeIMO.OfficeLength.Points(x + 10), global::OfficeIMO.OfficeLength.Points(168), global::OfficeIMO.OfficeLength.Points(105), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(18));
                if (index < 2) canvas.Text("→", global::OfficeIMO.OfficeLength.Points(x + 135), global::OfficeIMO.OfficeLength.Points(166), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(35), global::OfficeIMO.OfficeLength.Points(24));
            }
        })), font).ToBytes();
        var diagramGold = new EvaluationGold(Fields: new[] { new EvaluationFieldGold("nextStep", OfficeAiFieldStatus.Present, "Archive") }, Status: null);
        var diagramRequest = new OfficeAiRequest { Operation = OfficeAiOperation.ExtractFields,
            Instruction = "Follow the directed flow. What step immediately follows Review? Use the exact label as nextStep.",
            Fields = new[] { new OfficeAiFieldDefinition("nextStep") } };
        string reserve = "Project Birch requires 64 units. " + string.Concat(Enumerable.Repeat("Routine project notes contain no additional quantities. ", 450))
            + "\n\nProject Cedar requires 29 units. " + string.Concat(Enumerable.Repeat("Routine scheduling notes contain no additional deadlines. ", 450))
            + "\n\nBoth projects start on 2034-06-12.";
        var reserveGold = new EvaluationGold(FactMarkers: new[] { "64", "29", "2034-06-12" }, RequireSynthesis: true);
        return new[] {
            Case("columns-native-reserve", ".pdf", columns, fields, false, columnGold),
            Case("columns-scan-reserve", ".png", Raster(columns), fields, true, columnGold with { Status = null }),
            Case("diagram-scan-reserve", ".png", Raster(diagram), diagramRequest, true, diagramGold),
            Case("long-summary-reserve", ".txt", Encoding.UTF8.GetBytes(reserve), new() {
                Operation = OfficeAiOperation.Summarize, Instruction = "Summarize the required units for both projects and their common start date.",
                Limits = new() { MaxRequests = 32, Timeout = TimeSpan.FromMinutes(8) }
            }, false, reserveGold)
        };
    }

    private static EvaluationCase Case(string id, string extension, byte[] source, OfficeAiRequest request, bool images, EvaluationGold gold) =>
        new(id, extension, source, request, images, "Independent layout/reserve gold; inspect source and citations", gold) { Split = "heldout" };
    private static byte[] Raster(byte[] pdf) {
        using var document = PdfDocument.Load(new MemoryStream(pdf));
        return document.ExportImages(new PdfImageExportOptions {
            Format = OfficeImageExportFormat.Png, TargetDpi = 120, MaximumOutputCount = 1, MaximumRasterPixels = 5_000_000
        }).Single().Bytes;
    }
}
