using System;
using System.Collections.Generic;
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
            OfficeDrawingElement[] elements = EnumerateElements(drawing).ToArray();
            int shapes = elements.OfType<OfficeDrawingShape>().Count();
            int textRuns = elements.OfType<OfficeDrawingText>().Count();
            writer.WriteLine($"| {scenario.Name} | {scenario.Svg.Length} | {elements.Length} | {shapes} | {textRuns} | {unsupported} |");
        }
    }

    private static IEnumerable<OfficeDrawingElement> EnumerateElements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            yield return element;
            if (element is OfficeDrawingGroup group) {
                foreach (OfficeDrawingElement child in EnumerateElements(group.Drawing)) yield return child;
            } else if (element is OfficeDrawingEffectGroup effect) {
                foreach (OfficeDrawingElement child in EnumerateElements(effect.Drawing)) yield return child;
                if (effect.SoftMask != null) {
                    foreach (OfficeDrawingElement child in EnumerateElements(effect.SoftMask.Drawing)) yield return child;
                }
            }
        }
    }
}
