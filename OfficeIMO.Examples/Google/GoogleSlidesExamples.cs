using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Google {
    internal static class GoogleSlidesExamples {
        public static void Example_Plan(string folderPath) {
            Console.WriteLine("[*] Google Slides - Building translation plan");
            string filePath = Path.Combine(folderPath, "GoogleSlides.Plan.pptx");
            using PowerPointDocument deck = PowerPointDocument.Create(filePath);
            PowerPointSlide slide = deck.AddSlide();
            slide.AddTextBox("OfficeIMO Google Slides plan",global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(80));
            slide.Notes.Text = "This note remains editable in Google Slides.";
            deck.Save();
            GoogleSlidesTranslationPlan plan = deck.BuildGoogleSlidesPlan(new GoogleSlidesSaveOptions { Title = "OfficeIMO Google Slides plan" });
            foreach (TranslationNotice notice in plan.Report.Notices) Console.WriteLine($"  {notice.Severity}: {notice.Feature} - {notice.Message}");
        }

        public static async Task Example_ExportAsync(string folderPath) {
            Console.WriteLine("[*] Google Slides - Exporting presentation");
            GoogleWorkspaceSession? session = GoogleWorkspaceExampleSession.TryCreateSession();
            if (session == null) { GoogleWorkspaceExampleSession.PrintMissingTokenMessage(); return; }
            string filePath = Path.Combine(folderPath, "GoogleSlides.Export.pptx");
            using PowerPointDocument deck = PowerPointDocument.Create(filePath);
            deck.AddSlide().AddTextBox("OfficeIMO Google Slides export",global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(80));
            deck.Save();
            try {
                GooglePresentationReference result = await deck.ExportToGoogleSlidesAsync(session, new GoogleSlidesSaveOptions {
                    Title = "OfficeIMO Google Slides export",
                    Location = new GoogleDriveFileLocation { FolderId = session.Options.DefaultFolderId, DriveId = session.Options.DefaultDriveId },
                });
                Console.WriteLine($"  PresentationId: {result.PresentationId}");
                Console.WriteLine($"  WebViewLink: {result.WebViewLink}");
                foreach (TranslationNotice notice in result.Report.Notices) Console.WriteLine($"  {notice.Severity}: {notice.Feature} - {notice.Message}");
            } catch (GoogleWorkspaceExportException exception) {
                GoogleWorkspaceExampleSession.PrintExportFailure(exception);
            }
        }
    }
}
