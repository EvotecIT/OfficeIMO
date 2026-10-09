using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;

internal static partial class NativeKeynoteWriterSmoke {
    private const string RichAcceptedSha256 = "AA1A3834B31716024420F37B808AE0E8A164BD5BFDCC59719C0B17F5CC1F6654";

    private static void RunRichText() {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddRichText(new[] {
            new IWorkKeynoteTextRun("Plain ", "Helvetica Neue"),
            new IWorkKeynoteTextRun("bold", "Helvetica Neue", bold: true),
            new IWorkKeynoteTextRun(" ", "Helvetica Neue"),
            new IWorkKeynoteTextRun("italic", "Helvetica Neue", italic: true),
            new IWorkKeynoteTextRun(" ", "Helvetica Neue"),
            new IWorkKeynoteTextRun("underlined", "Helvetica Neue", underline: true),
            new IWorkKeynoteTextRun(".", "Helvetica Neue")
        }, 60, 100, 840, 150);
        document.AddSlide("E6F2FF").AddRichText(new[] {
            new IWorkKeynoteTextRun("Bold 😀", "Arial", 32, "003366", bold: true, italic: true,
                underline: true, strikethrough: true),
            new IWorkKeynoteTextRun(" plain", "Arial", 24),
            new IWorkKeynoteTextRun("\r\nItalic", "Times New Roman", 30, "008000", italic: true),
            new IWorkKeynoteTextRun("\nZażółć — café", "Arial", 24, "660066")
        }, 60, 80, 840, 250);
        document.AddSlide();
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == RichAcceptedSha256,
            "NativeAOT rich runs differ from the independently decoded and Apple-opened package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "NativeAOT rich runs are nondeterministic.");
        IWorkTextContent content = IWorkSourceDocument.Open(bytes).ReadKeynote().Slides[1].TextBoxes.Single().Content;
        Require(content.PlainText == "Bold 😀 plain\nItalic\nZażółć — café\n", "NativeAOT rich text lost Unicode or paragraph boundaries.");
        IWorkTextRun[] runs = content.Paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
        IWorkTextStyle first = runs.Single(run => run.Text == "Bold 😀").Style;
        IWorkTextStyle reset = runs.Single(run => run.Text == " plain").Style;
        Require(first.Bold == true && first.Italic == true && first.Underline == true && first.Strikethrough == true,
            "NativeAOT rich run lost decoration.");
        Require(reset.Bold == false && reset.Italic == false && reset.Underline == false && reset.Strikethrough == false
            && reset.FontSizePoints == 24 && reset.Color?.RgbHex == "000000", "NativeAOT plain run failed to reset formatting.");
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(bytes),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var saved = PowerPointPresentation.Load(output);
        Require(saved.Slides.Count == 3 && !saved.ValidateDocument().Any(), "NativeAOT rich-run PPTX did not save/reopen correctly.");
        PowerPointTextRun savedReset = saved.Slides[1].TextBoxes.Single().Paragraphs.SelectMany(paragraph => paragraph.Runs)
            .Single(run => run.Text == " plain");
        Require(!savedReset.Bold && !savedReset.Italic && !savedReset.Underline && !savedReset.Strikethrough,
            "NativeAOT saved PPTX inherited earlier character decoration.");
        using var unchanged = new MemoryStream(new byte[] { 9, 8, 7 });
        try {
            document.Save(unchanged, new IWorkKeynoteWriteOptions { MaximumTextRuns = 10 });
            throw new InvalidOperationException("NativeAOT character-run budget was not enforced.");
        } catch (InvalidDataException) { }
        Require(unchanged.ToArray().SequenceEqual(new byte[] { 9, 8, 7 }), "NativeAOT run-budget failure changed the destination.");
        Console.WriteLine($"PASS | native Keynote rich runs | {RichAcceptedSha256} | {bytes.Length} bytes | UTF-16 ranges, decoration resets, deterministic creation, saved/reopened PPTX and run limits");
    }
}
