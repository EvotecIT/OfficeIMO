using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PowerPointInheritedLinkOwnerTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void InheritedSoundReplacementSurvivesSaveReopen(bool master, bool table, bool mouseOver) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointShape inherited = CreateSoundConsumer(presentation, master, table);
        OpenXmlPart owner = Owner(presentation.Slides[0], master);
        byte[] original = SoundWave(), replacement = SoundWave();
        replacement[replacement.Length - 1] = 0x70;
        SetSound(inherited, original);
        SetSound(inherited, replacement);
        Assert.Single(owner.DataPartReferenceRelationships.OfType<AudioReferenceRelationship>());
        Assert.Single(presentation.OpenXmlDocument.DataParts);
        Assert.Empty(presentation.Slides[0].SlidePart.DataPartReferenceRelationships);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        PowerPointShape restored = Assert.Single(reopened.Slides[0].GetInheritedShapesForExport(), shape => shape.Name == inherited.Name);
        if (restored is PowerPointTable restoredTable) {
            PowerPointTextRun run = SoundRun(restoredTable);
            Assert.Equal(replacement, mouseOver ? run.GetMouseOverSoundBytes() : run.GetClickSoundBytes());
            if (mouseOver) run.ClearMouseOverSound(); else run.ClearClickSound();
        } else {
            Assert.Equal(replacement, mouseOver ? restored.GetMouseOverSoundBytes() : restored.GetClickSoundBytes());
            if (mouseOver) restored.ClearMouseOverSound(); else restored.ClearClickSound();
        }
        Assert.Empty(reopened.OpenXmlDocument.DataParts);
        Assert.Empty(reopened.ValidateDocument());

        void SetSound(PowerPointShape shape, byte[] bytes) {
            using var stream = new MemoryStream(bytes, writable: false);
            if (shape is PowerPointTable soundTable) {
                PowerPointTextRun run = SoundRun(soundTable);
                if (mouseOver) run.SetMouseOverSound(stream, "Sound"); else run.SetClickSound(stream, "Sound");
            } else if (mouseOver) shape.SetMouseOverSound(stream, "Sound");
            else shape.SetClickSound(stream, "Sound");
        }
    }

    [Theory]
    [InlineData(false, "row")]
    [InlineData(true, "row")]
    [InlineData(false, "column")]
    [InlineData(true, "column")]
    [InlineData(false, "merge")]
    [InlineData(true, "merge")]
    [InlineData(false, "clear")]
    [InlineData(true, "clear")]
    [InlineData(false, "text")]
    [InlineData(true, "text")]
    [InlineData(false, "paragraphs")]
    [InlineData(true, "paragraphs")]
    [InlineData(false, "markdown")]
    [InlineData(true, "markdown")]
    public void RemovingInheritedTextReleasesItsLastSound(bool master, string operation) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        bool table = operation is "row" or "column" or "merge";
        PowerPointShape inherited = CreateSoundConsumer(presentation, master, table);
        PowerPointTextRun run = table ? SoundRun((PowerPointTable)inherited) : ((PowerPointTextBox)inherited).Paragraphs[0].Runs[0];
        using (var audio = new MemoryStream(SoundWave(), writable: false)) run.SetClickSound(audio, "Discarded");
        switch (operation) {
            case "row": ((PowerPointTable)inherited).RemoveRow(1); break;
            case "column": ((PowerPointTable)inherited).RemoveColumn(1); break;
            case "merge": ((PowerPointTable)inherited).MergeCells(0, 0, 1, 1); break;
            case "clear": ((PowerPointTextBox)inherited).Clear(); break;
            case "text": ((PowerPointTextBox)inherited).Text = "Replacement"; break;
            case "paragraphs": ((PowerPointTextBox)inherited).SetParagraphs(["Replacement"]); break;
            case "markdown": ((PowerPointTextBox)inherited).SetMarkdown("**Replacement**"); break;
        }
        Assert.Empty(Owner(presentation.Slides[0], master).DataPartReferenceRelationships.OfType<AudioReferenceRelationship>());
        Assert.Empty(presentation.OpenXmlDocument.DataParts);
        Assert.Empty(presentation.ValidateDocument());
    }

    private static PowerPointShape CreateSoundConsumer(PowerPointDocument presentation, bool master, bool table) {
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointShape shape;
        if (table) {
            PowerPointTable value = slide.AddTable(2,2,global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(240),global::OfficeIMO.OfficeLength.Points(60));
            value.GetCell(1, 1).Text = "Sounded";
            shape = value;
        } else {
            PowerPointTextBox box = slide.AddTextBox("Sounded",global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(240),global::OfficeIMO.OfficeLength.Points(60));
            box.PlaceholderType = null;
            shape = box;
        }
        shape.Element.Remove(); Tree(Owner(slide, master)).Append(shape.Element);
        return Assert.Single(slide.GetInheritedShapesForExport(), inherited => inherited.Name == shape.Name);
    }

    private static PowerPointTextRun SoundRun(PowerPointTable table) => table.GetCell(1, 1).Paragraphs[0].Runs.Single(run => run.Text == "Sounded");

    private static byte[] SoundWave() => [82, 73, 70, 70, 40, 0, 0, 0, 87, 65, 86, 69,
        102, 109, 116, 32, 16, 0, 0, 0, 1, 0, 1, 0, 0x40, 0x1F, 0, 0,
        0x40, 0x1F, 0, 0, 1, 0, 8, 0, 100, 97, 116, 97, 4, 0, 0, 0, 0x80, 0x90, 0x70, 0x80];
}
