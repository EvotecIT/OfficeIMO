using OfficeIMO.PowerPoint;

namespace OfficeIMO.ConversionConsistency;

internal static partial class FixtureCorpus {
    private static void AddGroupedPowerPoint(string repository, string output, string family, List<ConsistencyCase> cases) {
        using var presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(600, 360);
        var variants = new[] {
            (ScaleX: 0.5D, ScaleY: 0.5D, Rotation: 0D, FlipX: false, FlipY: false, Nested: false),
            (ScaleX: 2D, ScaleY: 2D, Rotation: 0D, FlipX: false, FlipY: false, Nested: false),
            (ScaleX: 0.5D, ScaleY: 2D, Rotation: 0D, FlipX: false, FlipY: false, Nested: true),
            (ScaleX: 1D, ScaleY: 1D, Rotation: 30D, FlipX: false, FlipY: false, Nested: false),
            (ScaleX: 1D, ScaleY: 1D, Rotation: 0D, FlipX: true, FlipY: false, Nested: false),
            (ScaleX: 1D, ScaleY: 1D, Rotation: 0D, FlipX: false, FlipY: true, Nested: false)
        };
        for (int index = 0; index < variants.Length; index++) {
            var variant = variants[index];
            var slide = presentation.AddSlide();
            var box = slide.AddTextBoxPoints("GROUP" + (index + 1), 60, 70, 200, 60);
            box.FontName = family;
            box.FontSize = 14;
            box.Paragraphs[0].Runs[0].Underline = true;
            var bar = slide.AddRectanglePoints(60, 140, 200, 10);
            bar.FillColor = "10B981";
            var group = slide.GroupShapes(new PowerPointShape[] { box, bar });
            if (variant.Nested) group = slide.GroupShapes(new PowerPointShape[] {
                group, slide.AddRectanglePoints(60, 160, 200, 10)
            });
            group.WidthPoints *= variant.ScaleX;
            group.HeightPoints *= variant.ScaleY;
            group.Rotation = variant.Rotation;
            group.HorizontalFlip = variant.FlipX;
            group.VerticalFlip = variant.FlipY;
        }
        string source = Path.Combine(output, "grouped-text.pptx");
        presentation.Save(source);
        cases.Add(new ConsistencyCase {
            Id = "native-pptx-groups", Format = "pptx",
            Source = Path.GetRelativePath(repository, source).Replace('\\', '/'),
            Evidence = "Grouped searchable text with reduced, enlarged, nested nonuniform, rotated, and mirrored frames. Native PNG, browser-rendered SVG, and independently rasterized PDF must agree.",
            Pages = Enumerable.Range(1, variants.Length).Select(number => new PageExpectation {
                Width = 800, Height = 480, Text = new() { "GROUP" + number }
            }).ToList()
        });
    }
}
