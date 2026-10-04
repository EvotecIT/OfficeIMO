using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;

Require(!System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported
    && !System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeCompiled,
    "The iWork conversion host must execute as NativeAOT.");
var fixtures = new[] {
    ("simple.pages", "5AEE6D03277D2DB2104F593E64AFE081DEC539F0117B97124B6F99158124C93E"),
    ("simple.numbers", "D0B00D9CAE5985CCCAA3B2FB251FAE92EB0E38360FB4B5DF8B4350EB658F752B"),
    ("simple.key", "BA95755DF82CEB0CA834E1E03E2777C34FAD906320D8336B4F3FEFC6B48607EB"),
    ("formulas.numbers", "DD85BAD68898CE5B065F277C0B9BE1F3C32D696E3BAA6B09D3614BBD35A5249F")
};
foreach (var (name, hash) in fixtures) {
    string path = Path.Combine(AppContext.BaseDirectory, name);
    byte[] bytes = File.ReadAllBytes(path);
    Require(Convert.ToHexString(SHA256.HashData(bytes)) == hash, name + " fixture changed.");
    using var input = new MemoryStream(bytes, writable: false);
    using var saved = new MemoryStream();
    if (name.EndsWith(".pages", StringComparison.Ordinal)) {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(input);
        result.RequireCompleteEditableReconstruction();
        Require(!result.IsVisualFallback, "Pages unexpectedly used visual fallback.");
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Require(reopened.Paragraphs.Any(p => p.Text == "hello pages"), "Saved DOCX lost Pages body text.");
        Require(result.Report.PreservedRecords.Count > 0, "Pages lost source evidence.");
    } else if (name.EndsWith(".key", StringComparison.Ordinal)) {
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(input);
        result.RequireCompleteEditableReconstruction();
        Require(!result.IsVisualFallback, "Keynote unexpectedly used visual fallback.");
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Require(reopened.Slides.Count == 2 && reopened.Slides[0].TextBoxes.Any(),
            "Saved PPTX lost editable slides or text.");
        Require(reopened.Slides[0].TextBoxes.Any(t => t.Text.Contains("hello keynote", StringComparison.Ordinal))
            && reopened.Slides[0].TextBoxes.Any(t => t.Text.Contains("first bullet", StringComparison.Ordinal))
            && reopened.Slides[1].TextBoxes.Any(t => t.Text.Contains("second slide", StringComparison.Ordinal)),
            "Saved PPTX lost native titles or body text.");
        Require(reopened.Slides.Any(s => s.GetSpeakerNotesText().Contains("note text here", StringComparison.Ordinal)),
            "Saved PPTX lost presenter notes.");
        Require(reopened.Slides.All(s => s.BackgroundColor == "FFFFFF"), "Saved PPTX lost qualified backgrounds.");
        Require(!reopened.ValidateDocument().Any(), "Saved PPTX failed validation.");
        Require(result.Report.PreservedRecords.Count > 0, "Keynote lost source evidence.");
    } else {
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(input);
        result.RequireCompleteEditableReconstruction();
        Require(!result.IsVisualFallback, "Numbers unexpectedly used visual fallback.");
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = reopened.Sheets[0];
        if (name == "simple.numbers") {
            Require(sheet.CellAt(1, 1).GetValue<string>() == "a" && sheet.CellAt(2, 2).GetValue<double>() == 2
                && sheet.CellAt(3, 3).GetValue<string>() == "Z", "Saved XLSX lost typed native cells.");
        } else {
            Require(sheet.GetFormulaText(2, 2) == "A1+A2" && sheet.CellAt(2, 2).GetValue<double>() == 3,
                "Saved XLSX lost the native formula or cache.");
            sheet.CellValue(1, 1, 10d);
            Require(reopened.Calculate() > 0 && sheet.CellAt(2, 2).GetValue<double>() == 12,
                "NativeAOT XLSX recalculation did not follow the edited operand.");
        }
        Require(!reopened.ValidateOpenXml().Any(), "Saved XLSX failed validation.");
        Require(result.Report.PreservedRecords.Count > 0, "Numbers lost source evidence.");
    }
    Require(input.CanRead, "Conversion closed the caller-owned stream.");
    Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == hash, "Conversion changed source bytes.");
    Console.WriteLine($"PASS | {name} | {hash} | complete editable conversion, save/reopen, source preservation");
}
Console.WriteLine("PASS | bounded iWork NativeAOT destination conversions | "
    + System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}
