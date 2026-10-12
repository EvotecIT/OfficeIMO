using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointSpeakerNotesDiagnosticsTests {
    [Theory]
    [InlineData(PowerPointPdfPageLayout.Slides, false)]
    [InlineData(PowerPointPdfPageLayout.Slides, true)]
    [InlineData(PowerPointPdfPageLayout.NotesPages, false)]
    [InlineData(PowerPointPdfPageLayout.NotesPages, true)]
    [InlineData(PowerPointPdfPageLayout.Handouts, false)]
    [InlineData(PowerPointPdfPageLayout.Handouts, true)]
    public void Pdf_reports_present_speaker_notes_when_the_selected_layout_omits_them(
        PowerPointPdfPageLayout layout, bool includeNotes) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBoxPoints("Visible slide", 20, 20, 250, 40);
        slide.Notes.Text = "PRESENTER-NOTE";
        var options = new PowerPointToPdfOptions { PageLayout = layout, IncludeSpeakerNotes = includeNotes };

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        string text = PdfCore.PdfReadDocument.Open(result.Value.ToBytes()).ExtractText();
        bool omitted = layout == PowerPointPdfPageLayout.Slides || !includeNotes;
        if (omitted) {
            var warning = Assert.Single(result.Warnings, item => item.Code == "speaker-notes-omitted");
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.Equal("Slide 1", warning.Source);
            Assert.DoesNotContain("PRESENTER-NOTE", text);
        } else {
            Assert.DoesNotContain(result.Warnings, item => item.Code == "speaker-notes-omitted");
            Assert.Contains("PRESENTER-NOTE", text);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pdf_note_omission_diagnostics_follow_the_selected_slides(bool includeHidden) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBoxPoints("Visible slide", 20, 20, 250, 40);
        PowerPointSlide hidden = presentation.AddSlide();
        hidden.Hidden = true;
        hidden.Notes.Text = "Hidden presenter note";

        var result = presentation.ToPdfDocumentResult(new PowerPointToPdfOptions { IncludeHiddenSlides = includeHidden });

        if (includeHidden)
            Assert.Equal("Slide 2", Assert.Single(result.Warnings, item => item.Code == "speaker-notes-omitted").Source);
        else Assert.DoesNotContain(result.Warnings, item => item.Code == "speaker-notes-omitted");
    }
}
