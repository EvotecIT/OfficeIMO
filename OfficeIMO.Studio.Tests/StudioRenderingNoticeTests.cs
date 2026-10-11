using System.Globalization;
using System.Text;
using Avalonia;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioRenderingNoticeTests {
    [Fact]
    public async Task Skipped_content_raises_one_document_notice_while_routine_font_substitution_stays_on_the_page() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services) { Width = 1280, Height = 820 };
            try {
                Directory.CreateDirectory(services.Paths.Root);
                string source = Path.Combine(services.Paths.Root, "rendering-notes.pdf");
                File.WriteAllBytes(source, BuildTwoPagePdf(
                    "BT /F1 12 Tf 20 150 Td (Unembedded Helvetica) Tj ET",
                    "0 0 1 rg 20 20 80 80 re f 1 2 FuturePaint"));
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                MainWindowViewModel document = window.ViewModel;
                foreach (PdfPageViewModel page in document.Pages) page.AttachToViewport();
                await WaitUntilAsync(() => document.Pages.All(page => !page.IsRendering && page.HasDiagnostics));

                // An unembedded standard font is drawn with an installed font, as in every viewer: a page note only.
                Assert.Equal(PdfPageRenderingNote.SubstituteFont, document.Pages[0].RenderingNote);
                Assert.Equal(services.Localizer.Get("PdfPage.FontSubstitution"), document.Pages[0].DiagnosticsExplanation);
                // Skipped content can change what the reader sees, so the document says so once.
                Assert.Equal(PdfPageRenderingNote.ContentOmitted, document.Pages[1].RenderingNote);
                Assert.True(document.HasRenderingNotice);
                Assert.Equal(services.Localizer.Format("Notices.RenderingLimitation", 2), document.RenderingNotice);
                Assert.Equal(services.Localizer.Format("Notices.ShowPage", 2), document.RenderingNoticeAction);

                // The notice belongs to the document view, leads to the page, and can be put away for this document.
                document.ShowJobsCommand.Execute(null);
                Assert.False(document.HasRenderingNotice);
                document.ShowRenderingNoticePageCommand.Execute(null);
                Assert.True(document.IsPdfWorkspaceMode);
                Assert.Same(document.Pages[1], document.SelectedPage);
                document.DismissRenderingNoticeCommand.Execute(null);
                Assert.False(document.HasRenderingNotice);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public void Notes_are_classified_by_the_renderer_support_level() {
        Assert.Equal(PdfPageRenderingNote.None, PdfPageRenderingNotes.Classify(["Avalonia vector fallback: gradients require the OfficeIMO raster renderer."]));
        Assert.Equal(PdfPageRenderingNote.Simplified, PdfPageRenderingNotes.Classify(["render.operator.flatness-simplified: The flatness value is accepted."]));
        Assert.Equal(PdfPageRenderingNote.EmbeddedFontReplaced, PdfPageRenderingNotes.Classify([
            "render.resource.font-substitution: Mapped. Subject: F1.",
            "render.resource.font-cff-substitution: CFF outlines are not executed. Subject: F2."]));
        Assert.Equal(PdfPageRenderingNote.Simplified, PdfPageRenderingNotes.Classify(["render.resource.annotation-appearance-missing"]));
        Assert.Equal(PdfPageRenderingNote.ContentOmitted, PdfPageRenderingNotes.Classify(["render.operator.unsupported"]));
        Assert.False(PdfPageRenderingNotes.AffectsAppearance(PdfPageRenderingNote.SubstituteFont));
        Assert.True(PdfPageRenderingNotes.AffectsAppearance(PdfPageRenderingNote.EmbeddedFontReplaced));
    }

    private static byte[] BuildTwoPagePdf(string first, string second) {
        string Length(string content) => Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture);
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 200 200] /Resources << /Font << /F1 7 0 R >> >> >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>", "endobj",
            "4 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Length(first) + " >>", "stream", first, "endstream", "endobj",
            "6 0 obj", "<< /Length " + Length(second) + " >>", "stream", second, "endstream", "endobj",
            "7 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }));
    }

    private static async Task WaitUntilAsync(Func<bool> condition) {
        for (int attempt = 0; attempt < 200 && !condition(); attempt++) {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(20);
        }
        Assert.True(condition());
    }
}
