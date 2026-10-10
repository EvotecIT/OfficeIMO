using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    private static void WriteCellFontEvidence(WordDocumentVisualSnapshot snapshot, OfficeImageExportResult image,
        bool useShaper) {
        string? directory = Environment.GetEnvironmentVariable("OFFICE_CONSUMER_FONT_EVIDENCE");
        if (string.IsNullOrWhiteSpace(directory)) return;
        if (!Path.IsPathRooted(directory)) throw new InvalidOperationException("Font evidence requires an absolute output directory.");
        Directory.CreateDirectory(directory);
        string stem = Path.Combine(directory, useShaper ? "cell-shaper" : "cell-font");
        File.WriteAllBytes(stem + ".svg", image.Bytes);
        var evidence = new {
            Text = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Select(text => new {
                text.Text, text.X, text.Y, text.Width, text.Height, text.LineHeight,
                Font = text.Font, text.Padding, text.WrapText, text.ShrinkToFit
            }).ToArray(),
            RichText = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().Select(text => new {
                text.PlainText, text.X, text.Y, text.Width, text.Height, text.LineHeight,
                text.Padding, text.WrapText, text.ShrinkToFit, text.Runs
            }).ToArray(),
            Shapes = snapshot.Drawing.Shapes.Select(shape => new {
                shape.X, shape.Y, shape.Shape.Width, shape.Shape.Height, shape.Shape.Kind
            }).ToArray(),
            snapshot.Diagnostics
        };
        File.WriteAllText(stem + ".json", JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
    }
}
