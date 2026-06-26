using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word.Html;

byte[] onePixelPng = Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVEhLY1BNfv2flpgBXYDaeBhaILCzkSKMbt6oBRgY3bxRCzAwunmjFmBgdPNGLcDA6OaNWoCB0c3DsIDaeNQCghgAFxBXzP1LTe4AAAAASUVORK5CYII=");

string outputRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "artifacts"));
Directory.CreateDirectory(outputRoot);

string statusImageDataUri = "data:image/png;base64," + Convert.ToBase64String(onePixelPng);
string wordSourceHtml = $$"""
    <!doctype html>
    <html lang="en">
    <head>
        <title>Word Roundtrip Proof</title>
        <style>
            body { font-family: Aptos, Calibri, sans-serif; color: #1f2937; }
            article { max-width: 860px; margin: 0 auto; }
            .lead { color: #2563eb; font-size: 15pt; }
            table.report { width: 100%; border-collapse: collapse; margin-top: 18px; }
            th { background: #dbeafe; color: #0f172a; }
            th, td { border: 1px solid #94a3b8; padding: 7pt; }
            figure { margin: 20px 0; }
            figcaption { color: #475569; font-size: 10pt; }
        </style>
    </head>
    <body>
        <article id="word-proof">
            <h1>Word Roundtrip Proof</h1>
            <p class="lead">HTML imports to Word, validates as DOCX, and exports back to styled HTML with shared manifest evidence.</p>
            <ul>
                <li><input type="checkbox" checked> Shared profile contract recorded</li>
                <li><input type="checkbox"> Print-review proof still tracked as a known future gap</li>
            </ul>
            <table class="report">
                <thead>
                    <tr><th>Feature</th><th>Evidence</th><th>Status</th></tr>
                </thead>
                <tbody>
                    <tr><td>Tables</td><td>thead/tbody/tfoot roundtrip</td><td>Preserved</td></tr>
                    <tr><td>Forms</td><td>checkbox/select controls</td><td>Preserved</td></tr>
                    <tr><td>Images</td><td>embedded data URI image</td><td>Preserved</td></tr>
                </tbody>
                <tfoot>
                    <tr><td colspan="2">Manifest proof</td><td>Shared JSON</td></tr>
                </tfoot>
            </table>
            <figure>
                <img src="{{statusImageDataUri}}" alt="Inline proof badge" width="48" height="48">
                <figcaption>Inline image survives through the Word HTML adapter path as inspectable HTML evidence.</figcaption>
            </figure>
            <label>Owner <input name="owner" value="OfficeIMO"></label>
            <label>Status <select name="status"><option>Draft</option><option selected>Validated</option></select></label>
            <!-- visible only as HtmlCommentSkipped diagnostic in the shared manifest -->
        </article>
    </body>
    </html>
    """;

wordSourceHtml.SaveHtmlCapabilityGallery(outputRoot, new WordHtmlCapabilityGalleryOptions {
    ScenarioId = "word-roundtrip",
    Title = "Word Roundtrip HTML Proof"
});

using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
ExcelSheet sheet = workbook.AddWorkSheet("Quarterly Sales");
sheet.CellValue(1, 1, "Region");
sheet.CellValue(1, 2, "Q1");
sheet.CellValue(1, 3, "Q2");
sheet.CellValue(1, 4, "Status");
sheet.CellValue(2, 1, "North");
sheet.CellValue(2, 2, 124000);
sheet.CellValue(2, 3, 139500);
sheet.CellValue(2, 4, "Ahead");
sheet.CellValue(3, 1, "South");
sheet.CellValue(3, 2, 98000);
sheet.CellValue(3, 3, 105300);
sheet.CellValue(3, 4, "Stable");
sheet.CellValue(4, 1, "West");
sheet.CellValue(4, 2, 87500);
sheet.CellValue(4, 3, 112750);
sheet.CellValue(4, 4, "Recovering");
sheet.CellBackground(1, 1, "#DBEAFE");
sheet.CellBackground(1, 2, "#DBEAFE");
sheet.CellBackground(1, 3, "#DBEAFE");
sheet.CellBackground(1, 4, "#DBEAFE");
sheet.CellBold(1, 1);
sheet.CellBold(1, 2);
sheet.CellBold(1, 3);
sheet.CellBold(1, 4);
sheet.CellFormula(5, 2, "SUM(B2:B4)");
sheet.CellValue(5, 1, "Total Q1");
sheet.SetComment(2, 4, "Status reviewed during forecast meeting", "OfficeIMO");
sheet.AddChartFromRange("A1:C4", row: 1, column: 6, widthPixels: 360, heightPixels: 220, type: ExcelChartType.ColumnClustered, title: "Regional Revenue Trend");
sheet.AddImage(5, 4, onePixelPng, widthPixels: 56, heightPixels: 56, name: "Status Logo", altText: "Inline status image");

workbook.SaveHtmlCapabilityGallery(outputRoot, new ExcelHtmlCapabilityGalleryOptions {
    ScenarioId = "excel-rich",
    Title = "Excel Rich HTML Proof",
    Theme = OfficeHtmlDocumentThemeKind.Report
});
CopyArtifact(outputRoot, "excel-rich.semantic.html", "excel-semantic.html");
CopyArtifact(outputRoot, "excel-rich.visual.html", "excel-visual.html");

using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
PowerPointSlide slide = presentation.Slides[0];
slide.AddTitlePoints("HTML Roundtrip Plan", 48, 36, 520, 54);
slide.AddTextBoxPoints("Shared OfficeIMO.Html profiles keep semantic and positioned review lanes honest.", 64, 120, 520, 78);
using (var image = new MemoryStream(onePixelPng)) {
    PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 558, 40, 72, 72);
    picture.Name = "Status badge";
    picture.AltText = "Reusable renderer badge";
}

PowerPointTable table = slide.AddTablePoints(3, 2, 64, 230, 430, 120);
table.GetCell(0, 0).Text = "Lane";
table.GetCell(0, 1).Text = "Proof";
table.GetCell(1, 0).Text = "Semantic";
table.GetCell(1, 1).Text = "Text/table/notes";
table.GetCell(2, 0).Text = "Visual review";
table.GetCell(2, 1).Text = "Positioned geometry";
PowerPointChartData chartData = new(
    new[] { "Q1", "Q2", "Q3" },
    new[] {
        new PowerPointChartSeries("Actual", new[] { 12D, 18D, 24D }),
        new PowerPointChartSeries("Target", new[] { 15D, 20D, 22D })
    });
slide.AddChartPoints(chartData, 520, 180, 240, 150).SetTitle("Pipeline Trend");
slide.Notes.Text = "Presenter note: prove the adapter outputs readable HTML and explicit visual-review boundaries.";

presentation.SaveHtmlCapabilityGallery(outputRoot, new PowerPointHtmlCapabilityGalleryOptions {
    ScenarioId = "powerpoint-rich",
    Title = "PowerPoint Rich HTML Proof",
    Theme = OfficeHtmlDocumentThemeKind.Report,
    IncludeNotes = true,
    IncludeTables = true
});
CopyArtifact(outputRoot, "powerpoint-rich.semantic.html", "powerpoint-semantic.html");
CopyArtifact(outputRoot, "powerpoint-rich.visual.html", "powerpoint-visual.html");

Console.WriteLine(outputRoot);

static void CopyArtifact(string directory, string sourceName, string targetName) {
    File.Copy(Path.Combine(directory, sourceName), Path.Combine(directory, targetName), overwrite: true);
}
