using OfficeIMO;
using OfficeIMO.PowerPoint;

internal static class HiddenDataSmoke {
    internal static void Run() {
        using var stream = new MemoryStream();
        using PowerPointDocument presentation = PowerPointDocument.Create(stream, new PowerPointCreateOptions());
        PowerPointSlide slide = presentation.AddSlide();
        slide.Notes.Text = "Private speaker note";
        OfficeHiddenDataReport report = presentation.RemoveHiddenData();
        if (!report.Changed || slide.HasSpeakerNotes) {
            throw new InvalidOperationException("PowerPoint hidden-data removal did not remove the speaker notes.");
        }
        Console.WriteLine("PASS | PowerPoint hidden-data removal");
    }
}
