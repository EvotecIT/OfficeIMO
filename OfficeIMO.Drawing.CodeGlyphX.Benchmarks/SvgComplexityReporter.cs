using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;

internal static class SvgComplexityReporter {
    internal static void Write(TextWriter writer) {
        writer.WriteLine("| Scenario | SVG bytes | Drawing elements | Shapes | Text runs | Unsupported |");
        writer.WriteLine("| --- | ---: | ---: | ---: | ---: | ---: |");
        foreach (string name in SvgScenarioFactory.Names) {
            SvgScenario scenario = SvgScenarioFactory.Create(name);
            if (!OfficeSvgDrawingReader.TryRead(scenario.Svg, out OfficeDrawing? drawing, out int unsupported) || drawing is null) {
                throw new InvalidOperationException($"Could not import complexity scenario '{name}'.");
            }
            int textRuns = drawing.Elements.OfType<OfficeDrawingText>().Count();
            writer.WriteLine($"| {scenario.Name} | {scenario.Svg.Length} | {drawing.Elements.Count} | {drawing.Shapes.Count} | {textRuns} | {unsupported} |");
        }
    }
}
