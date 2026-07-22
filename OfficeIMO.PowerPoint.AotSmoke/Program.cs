using OfficeIMO.PowerPoint;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".pptx");
try {
    using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
        presentation.AddSlide().AddTitle("OfficeIMO NativeAOT slide");
        presentation.Save();
    }

    using PowerPointPresentation reopened = PowerPointPresentation.Load(path);
    if (reopened.Slides.Count != 1) throw new InvalidOperationException("The PowerPoint round trip lost its slide.");

    Console.WriteLine("PASS | PowerPoint create, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
