using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".pptx");
try {
    using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
        var data = new OfficeChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeChartSeries("Revenue", new[] { 1d, 2d, 3d }) });

        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("OfficeIMO NativeAOT slide");
        slide.AddChart(OfficeChartKind.ColumnClustered, data);
        presentation.DuplicateSlide(0);
        presentation.Save();
    }

    using PowerPointPresentation reopened = PowerPointPresentation.Load(path);
    if (reopened.Slides.Count != 2 || reopened.Slides[0].Charts.Count() != 1 || reopened.Slides[1].Charts.Count() != 1) {
        throw new InvalidOperationException("The PowerPoint round trip lost its slide or cloned chart relationships.");
    }

    Console.WriteLine("PASS | PowerPoint chart create, duplicate, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
