using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

internal static class NativeKeynoteParagraphEditingSmoke {
    internal static void Run() {
        var source = Open("preservation.key"); var deck = source.ReadKeynote(); var edit = source.BeginEdit();
        string[] values = ["Edited PNG — image behind text", "Edited alpha PNG", "Edited JPEG"];
        for (int index = 0; index < values.Length; index++)
            edit.SetKeynoteTextBoxParagraphText(deck.Slides[index].TextBoxes[0].SourceIdentity!, 1, values[index]);
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        var output = edit.SaveBytesResult(options);
        Require(Convert.ToHexString(SHA256.HashData(output.GetPackageBytes())) == "B4B6A57FBCD4FC724FA34AEB78A8EA7702CC350439069278492E3332D26AD254"
            && output.Report.UpdatedRecords.Count == 3, "Keynote text edits differ from the independent/native qualification.");
        Check(IWorkSourceDocument.Open(output.GetPackageBytes()), "Edited JPEG");
        Check(Open("keynote-paragraph-edited.key"), "Native checked JPEG");
        var first = deck.Slides[0].TextBoxes[0]; byte[] before = output.GetPackageBytes();
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        try { edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, first.Content.Paragraphs[0].Text, cancellation.Token); throw new Exception("Keynote text edit ignored cancellation."); }
        catch (OperationCanceledException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Cancelled Keynote text edit changed pending values."); }
        try { edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, "Length change"); throw new Exception("Keynote text length change was accepted."); }
        catch (NotSupportedException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Rejected Keynote text edit changed pending values."); }
        foreach (var slide in deck.Slides) {
            var box = slide.TextBoxes[0];
            edit.SetKeynoteTextBoxParagraphText(box.SourceIdentity!, 1, box.Content.Paragraphs[0].Text);
        }
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0 && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(), "Restoring Keynote text did not restore original bytes.");
        Console.WriteLine("PASS | native Keynote paragraph editing | three exact UTF-8 ranges, fixed UTF-16 offsets, unchanged images/geometry, Apple edit/reopen, saved PPTX, cancellation, length rejection and restoration");
    }

    private static void Check(IWorkSourceDocument source, string finalTitle) {
        var deck = source.ReadKeynote();
        Require(deck.HasEditableContent && deck.Slides.Count == 3
            && deck.Slides[0].TextBoxes[0].Content.Paragraphs[0].Text == "Edited PNG — image behind text"
            && deck.Slides[1].TextBoxes[0].Content.Paragraphs[0].Text == "Edited alpha PNG"
            && deck.Slides[2].TextBoxes[0].Content.Paragraphs[0].Text == finalTitle
            && deck.Slides[1].TextBoxes[1].Content.Paragraphs[0].Runs.Count == 2, "Edited Keynote content or untouched rich runs differ.");
        Require(deck.Slides.All(slide => slide.Images.Count == 1)
            && deck.Slides[0].Images[0].Geometry!.WidthPoints == 480
            && deck.Slides[1].Images[0].Geometry!.WidthPoints == 480
            && deck.Slides[2].Images[0].Geometry!.WidthPoints == 600, "Keynote text edit changed image frames.");
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(source.BeginEdit().SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction(); using var bytes = new MemoryStream(); converted.Value.Save(bytes);
        using var xml = PresentationDocument.Open(new MemoryStream(bytes.ToArray()), false);
        var paragraphs = xml.PresentationPart!.SlideParts.SelectMany(part => part.Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            .Select(item => item.InnerText).ToArray();
        Require(new[] { "Edited PNG — image behind text", "Edited alpha PNG", finalTitle, "100% alpha" }
            .All(text => paragraphs.Any(paragraph => paragraph.Contains(text, StringComparison.Ordinal))), "Saved PPTX lost edited or untouched text.");
    }
    private static IWorkSourceDocument Open(string name) => IWorkSourceDocument.Open(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name)));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
