using OfficeIMO.Drawing;
using System.Security.Cryptography;

internal static class PackedMathContract {
    // This generated rectangular-outline font has independently inspected MATH
    // records (ManagedTextShapingTestAssets.MathGlyphs.cs). No system font or
    // project reference can hide a packaged defect.
    public static void Verify() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", "FixtureMath.ttf"));
        using (var sha = SHA256.Create()) {
            string hash = BitConverter.ToString(sha.ComputeHash(font)).Replace("-", "").ToLowerInvariant();
            Require(hash == "db8334b017a154cd4e80aa048e226dfb308c67fde671b97c1061f649fa203613",
                "The packed MATH fixture bytes changed.");
        }
        var options = new OfficeMathRenderOptions {
            Font = new OfficeFontInfo("Fixture Math", 20D), Padding = 0D
        };
        options.Fonts.Add("Fixture Math", font);

        OfficeDrawing display = OfficeMathRenderer.Render(OfficeMath.Operator("∑"), options);
        OfficeDrawingShape variant = display.Elements.OfType<OfficeDrawingShape>().Single();
        Near(variant.Shape.Width, 10D); Near(variant.Shape.Height, 44D);
        Require(display.Elements.OfType<OfficeDrawingText>().Single().Text == "∑", "Operator text was lost.");

        OfficeMathExpression content = OfficeMath.Fraction(
            OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Identifier("y")),
            OfficeMath.Fraction(OfficeMath.Identifier("y"), OfficeMath.Number("2")));
        OfficeDrawing fenced = OfficeMathRenderer.Render(OfficeMath.Delimited(content), options);
        OfficeDrawingShape[] fences = fenced.Elements.OfType<OfficeDrawingShape>()
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Path).ToArray();
        Require(fences.Length == 2, "Tall fences did not use glyph assemblies.");
        foreach (OfficeDrawingShape fence in fences) {
            Near(fence.Shape.Width, 10D);
            Require(fence.Shape.Height > 44D && fence.Y >= 0D &&
                fence.Y + fence.Shape.Height <= fenced.Height + .000001D, "Fence paint escaped its frame.");
        }
        Require(fenced.Elements.OfType<OfficeDrawingText>().Count(text => text.Text == "(") == 1 &&
            fenced.Elements.OfType<OfficeDrawingText>().Count(text => text.Text == ")") == 1,
            "Fence logical text was duplicated or lost.");

        OfficeDrawing radical = OfficeMathRenderer.Render(
            OfficeMath.Radical(OfficeMath.Identifier("x"), OfficeMath.Number("2")), options);
        Near(radical.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "2").Font.Size, 11D);
        Near(radical.Elements.OfType<OfficeDrawingShape>().Single(shape => shape.Shape.Kind == OfficeShapeKind.Line)
            .Shape.StrokeWidth, 1.36D);

        var scripts = OfficeMathRenderer.Render(OfficeMath.SubSuperscript(
            OfficeMath.Identifier("x"), OfficeMath.Identifier("i"), OfficeMath.Number("2")), options)
            .Elements.OfType<OfficeDrawingText>().ToDictionary(text => text.Text);
        Near(scripts["2"].X, scripts["x"].X + scripts["x"].Width + 2D);
        Near(scripts["i"].X, scripts["x"].X + scripts["x"].Width - 2D);

        var accents = OfficeMathRenderer.Render(OfficeMath.Accent(OfficeMath.Identifier("x"), "^"), options)
            .Elements.OfType<OfficeDrawingText>().ToDictionary(text => text.Text);
        Near(accents["x"].X + 7D, accents["^"].X + 2D);
        Near(accents["^"].Font.Size, 20D);
        Near(accents["^"].Y + accents["^"].Height, accents["x"].Y);
        OfficeDrawing wideAccent = OfficeMathRenderer.Render(OfficeMath.Accent(
            OfficeMath.Row(Enumerable.Range(0, 6).Select(_ => OfficeMath.Identifier("x")).ToArray()), "^"), options);
        OfficeDrawingShape accentAssembly = wideAccent.Elements.OfType<OfficeDrawingShape>().Single();
        Near(accentAssembly.Shape.Width, 60D); Near(accentAssembly.Shape.Height, 4D);
        Require(wideAccent.Elements.OfType<OfficeDrawingText>().Count(text => text.Text == "^") == 1,
            "Horizontal accent assembly duplicated its logical text.");

        options.UseFontMathMetrics = false; options.ScriptScale = .5D;
        var callerScripts = OfficeMathRenderer.Render(
            OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2")), options)
            .Elements.OfType<OfficeDrawingText>().ToDictionary(text => text.Text);
        Near(callerScripts["2"].X, callerScripts["x"].X + callerScripts["x"].Width);
        Near(OfficeMathRenderer.Render(OfficeMath.Accent(OfficeMath.Identifier("x"), "^"), options)
            .Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "^").Font.Size, 10D);
        OfficeDrawing fallback = OfficeMathRenderer.Render(OfficeMath.Operator("∑"),
            new OfficeMathRenderOptions { Font = new OfficeFontInfo("Fixture Math", 20D), Padding = 0D });
        Require(fallback.Elements.OfType<OfficeDrawingText>().Single().Text == "∑",
            "Missing-font fallback lost its logical operator.");

        var untouched = new OfficeDrawing(100D, 100D);
        using (var cancelled = new CancellationTokenSource()) {
            cancelled.Cancel();
            try {
                OfficeMathRenderer.AddToDrawing(untouched, content, 0D, 0D, options, cancelled.Token);
                throw new InvalidOperationException("Packed math ignored cancellation.");
            } catch (OperationCanceledException) {
                Require(untouched.Elements.Count == 0, "Cancelled math added partial paint.");
            }
        }
        Console.WriteLine("Packed MATH variants, assemblies, radical, placement, fallback and cancellation passed.");
    }

    private static void Near(double actual, double expected) =>
        Require(Math.Abs(actual - expected) <= .000001D, "Packed MATH geometry differed: " + actual + " vs " + expected);

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
