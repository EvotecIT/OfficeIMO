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
        var options = new PowerPointPdfExportOptions { PageLayout = layout, IncludeSpeakerNotes = includeNotes };

        PdfCore.PdfDocumentConversionResult result = presentation.ConvertToPdfWithEvidence(options);
        string text = PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText();
        bool omitted = layout == PowerPointPdfPageLayout.Slides || !includeNotes;
        if (omitted) {
            var warning = Assert.Single(result.Warnings, item => item.Code == "speaker-notes-omitted");
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.Equal("Slide 1", warning.Source);
            Assert.DoesNotContain("PRESENTER-NOTE", text);
        } else {
            Assert.DoesNotContain(result.Warnings, item => item.Code == "speaker-notes-omitted");
            Assert.DoesNotContain(result.Warnings, item => item.Code == "speaker-notes-overflow");
            Assert.Contains("PRESENTER-NOTE", text);
        }
    }

    [Theory]
    [InlineData(PowerPointPdfPageLayout.NotesPages, 6, 100)]
    [InlineData(PowerPointPdfPageLayout.Handouts, 3, 40)]
    [InlineData(PowerPointPdfPageLayout.Handouts, 6, 10)]
    public void Pdf_reports_enabled_notes_that_exceed_their_rendered_frame(
        PowerPointPdfPageLayout layout, int slidesPerPage, int lines) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBoxPoints("Visible slide", 20, 20, 250, 40);
        slide.Notes.Text = string.Join("\n", Enumerable.Range(1, lines).Select(index => $"PRESENTER-NOTE-{index:000}"));
        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions {
            PageLayout = layout, IncludeSpeakerNotes = true, HandoutSlidesPerPage = slidesPerPage
        });

        byte[] bytes = result.ToBytes();
        string? capture = Environment.GetEnvironmentVariable("OFFICEIMO_PPT_NOTES_VISUAL_OUTPUT");
        if (!string.IsNullOrWhiteSpace(capture)) {
            Directory.CreateDirectory(capture);
            File.WriteAllBytes(Path.Combine(capture, $"notes-{layout}-{slidesPerPage}.pdf"), bytes);
        }

        Assert.NotEmpty(PdfCore.PdfReadDocument.Open(bytes).Pages);
        var warning = Assert.Single(result.Warnings, item => item.Code == "speaker-notes-overflow");
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Equal("Slide 1", warning.Source);
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("PowerPointSpeakerNotes", warning.LayoutDiagnostic.Source);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
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

        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions { IncludeHiddenSlides = includeHidden });

        if (includeHidden)
            Assert.Equal("Slide 2", Assert.Single(result.Warnings, item => item.Code == "speaker-notes-omitted").Source);
        else Assert.DoesNotContain(result.Warnings, item => item.Code == "speaker-notes-omitted");
    }
}
