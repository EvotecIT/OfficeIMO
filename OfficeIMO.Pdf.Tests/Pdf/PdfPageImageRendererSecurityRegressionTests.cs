using System;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void Type3VisibilityGeometryBudgetBoundsRepeatedClipSeparationProofs() {
        PdfPageClipPath first = PdfPageClipPath.Rectangle(0D, 0D, 10D, 10D);
        PdfPageClipPath second = PdfPageClipPath.Rectangle(20D, 20D, 10D, 10D);
        var budget = new PdfReadPage.VisualGeometryBudget();
        bool provedSeparated = true;

        for (int index = 0; index < 5000 && provedSeparated; index++) {
            provedSeparated = first.CanProveNoPositiveAreaIntersection(second, budget);
        }

        Assert.True(budget.Exceeded);
        Assert.False(provedSeparated);
    }

    [Fact]
    public void RenderPage_ChargesType3CharProcAgainstContentNestingBudget() {
        string form = BuildStreamObject(5, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << /Font << /FType3 6 0 R >> >>", "BT /FType3 18 Tf 20 100 Td (A) Tj ET");
        string type3Font = "6 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 7 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyphA = BuildStreamObject(7, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("/Fm1 Do", "<< /XObject << /Fm1 5 0 R >> >>", form, type3Font, glyphA);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 1 }
        });

        PdfReadLimitException diagnosticException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).AssessRenderCompatibility(new PdfLoadOptions
            {
                Limits = new PdfReadLimits { MaxContentNestingDepth = 1 }
            }));
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, diagnosticException.Kind);
        Assert.Equal(1, diagnosticException.Limit);
        Assert.Equal(2, diagnosticException.Actual);
        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void RenderPage_AlignsAnnotationType3DepthBetweenDiagnosticsAndDrawing() {
        string annotation = "5 0 obj\n<< /Type /Annot /Subtype /Stamp /Rect [20 20 80 80] /AP << /N 6 0 R >> >>\nendobj";
        string appearance = BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 60 60] /Resources << /Font << /FType3 7 0 R >> >>", "BT /FType3 18 Tf (A) Tj ET");
        string type3Font = "7 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 8 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyphA = BuildStreamObject(8, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdfWithPageEntries("", "<< >>", "/Annots [5 0 R]", annotation, appearance, type3Font, glyphA);

        AssertAuxiliaryType3DepthMatchesDrawing(pdf);
    }

    [Fact]
    public void RenderPage_ChargesTilingPatternType3DepthInDiagnosticsAndDrawing() {
        string pattern = BuildStreamObject(5, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /FType3 6 0 R >> >>", "BT /FType3 8 Tf (A) Tj ET");
        string type3Font = "6 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 7 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyphA = BuildStreamObject(7, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("/Pattern cs /P1 scn 0 0 20 20 re f", "<< /Pattern << /P1 5 0 R >> >>", pattern, type3Font, glyphA);

        AssertAuxiliaryType3DepthIsBounded(pdf);
    }

    [Fact]
    public void RenderPage_AlignsSoftMaskType3DepthBetweenDiagnosticsAndDrawing() {
        string graphicsState = "5 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 6 0 R >> >>\nendobj";
        string softMask = BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FType3 7 0 R >> >>", "BT /FType3 8 Tf (A) Tj ET");
        string type3Font = "7 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 8 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyphA = BuildStreamObject(8, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("/GS1 gs 0 0 20 20 re f", "<< /ExtGState << /GS1 5 0 R >> >>", graphicsState, softMask, type3Font, glyphA);

        AssertAuxiliaryType3DepthMatchesDrawing(pdf);
    }

    [Fact]
    public void RenderCompatibility_BoundsNestedTilingPatternSurfaces() {
        string pattern1 = BuildStreamObject(5, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Pattern << /P2 6 0 R >> >>", "/Pattern cs /P2 scn 0 0 10 10 re f");
        string pattern2 = BuildStreamObject(6, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Pattern << /P3 7 0 R >> >>", "/Pattern cs /P3 scn 0 0 10 10 re f");
        string pattern3 = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >>", "0 0 10 10 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs /P1 scn 0 0 20 20 re f",
            "<< /Pattern << /P1 5 0 R >> >>",
            pattern1,
            pattern2,
            pattern3);

        AssertNestedAuxiliarySurfaceLimit(pdf);
    }

    [Fact]
    public void RenderCompatibility_BoundsNestedSoftMaskSurfaces() {
        string graphicsState1 = "5 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 6 0 R >> >>\nendobj";
        string softMask1 = BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency >> /Resources << /ExtGState << /GS2 7 0 R >> >>", "/GS2 gs 0 0 20 20 re f");
        string graphicsState2 = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string softMask2 = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency >> /Resources << /ExtGState << /GS3 9 0 R >> >>", "/GS3 gs 0 0 20 20 re f");
        string graphicsState3 = "9 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 10 0 R >> >>\nendobj";
        string softMask3 = BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency >> /Resources << >>", "0 0 20 20 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "/GS1 gs 0 0 20 20 re f",
            "<< /ExtGState << /GS1 5 0 R >> >>",
            graphicsState1,
            softMask1,
            graphicsState2,
            softMask2,
            graphicsState3,
            softMask3);

        AssertNestedAuxiliarySurfaceLimit(pdf);
    }

    private static void AssertNestedAuxiliarySurfaceLimit(byte[] pdf) {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).AssessRenderCompatibility(new PdfLoadOptions
            {
                Limits = new PdfReadLimits { MaxContentNestingDepth = 1 }
            }));

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    private static void AssertAuxiliaryType3DepthMatchesDrawing(byte[] pdf) {
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 1 }
        };

        _ = PdfDocument.Load(pdf).AssessRenderCompatibility(options);
        OfficeDrawing drawing = PdfReadDocument.Open(pdf, options).Pages[0].ToDrawing();

        Assert.NotEmpty(drawing.Elements);
    }

    private static void AssertAuxiliaryType3DepthIsBounded(byte[] pdf) {
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 1 }
        };

        PdfReadLimitException diagnosticException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).AssessRenderCompatibility(options));
        PdfReadLimitException drawingException = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, diagnosticException.Kind);
        Assert.Equal(2, diagnosticException.Actual);
        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, drawingException.Kind);
        Assert.Equal(2, drawingException.Actual);
    }

    [Fact]
    public void RenderPage_BoundsPendingTextClippingPaths() {
        string textShows = string.Concat(Enumerable.Repeat(
            "(A) Tj ",
            PdfPageClipPath.MaximumPendingTextClippingPaths + 1));
        byte[] pdf = BuildSingleStreamPdf(
            "BT /F1 12 Tf 4 Tr " + textShows + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.True(
            exception.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected limit kind: {exception.Kind}.");
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPage_BoundsAggregateTextClippingPathsAcrossTextObjects() {
        int firstCount = PdfPageClipPath.MaximumPendingTextClippingPaths / 2;
        string first = string.Concat(Enumerable.Repeat("(A) Tj ", firstCount));
        string second = string.Concat(Enumerable.Repeat("(A) Tj ", firstCount + 1));
        byte[] pdf = BuildSingleStreamPdf(
            "BT /F1 12 Tf 4 Tr " + first + "ET BT /F1 12 Tf 4 Tr " + second + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.True(
            exception.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected limit kind: {exception.Kind}.");
        Assert.True(exception.Actual > exception.Limit);

        PdfReadLimitException diagnosticException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).AssessRenderCompatibility());
        Assert.True(
            diagnosticException.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected diagnostic limit kind: {diagnosticException.Kind}.");
        Assert.True(diagnosticException.Actual > diagnosticException.Limit);
    }

    [Fact]
    public void RenderPage_DoesNotDoubleChargeTextClipsAcrossParserPasses() {
        int textClipCount = PdfPageClipPath.MaximumPendingTextClippingPaths / 2 + 1;
        string textShows = string.Concat(Enumerable.Repeat("(A) Tj ", textClipCount));
        byte[] pdf = BuildSingleStreamPdf(
            "BT /F1 12 Tf 4 Tr " + textShows + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        OfficeDrawing drawing = document.Pages[0].ToDrawing();

        Assert.NotNull(drawing);
    }

    [Fact]
    public void RenderPage_BoundsAggregateTextClipIntersectionWork() {
        const int runsPerObject = 1001;
        string runs = string.Concat(Enumerable.Repeat("(A) Tj ", runsPerObject));
        byte[] pdf = BuildSingleStreamPdf(
            "BT /F1 12 Tf 4 Tr " + runs + "ET BT /F1 12 Tf 4 Tr " + runs + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPage_ChargesVerticesDuringTextClipIntersections() {
        string activeClip = "0 -20 m " + string.Concat(Enumerable.Range(1, 300)
            .Select(index => index + " " + (index % 2 == 0 ? -20 : 100) + " l ")) + "h W n ";
        string textShows = string.Concat(Enumerable.Repeat("(A) Tj ", 1000));
        byte[] pdf = BuildSingleStreamPdf(
            activeClip + "BT /F1 12 Tf 4 Tr " + textShows + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPage_ChargesPathCommandsDuringRectangleTextClipIntersections() {
        var commands = new List<OfficePathCommand> {
            OfficePathCommand.MoveTo(new OfficePoint(0D, -20D))
        };
        commands.AddRange(Enumerable.Range(1, 300).Select(index =>
            OfficePathCommand.LineTo(new OfficePoint(index, index % 2 == 0 ? -20D : 100D))));
        commands.Add(OfficePathCommand.Close());
        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath path));
        path = path.AsTextClippingPath();
        PdfPageClipPath rectangle = PdfPageClipPath.Rectangle(0D, -20D, 300D, 120D);
        var budget = new PdfTextClippingBudget();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => {
            for (int index = 0; index < 3400; index++) {
                budget.ResolveActiveClip(path, rectangle);
            }
        });

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPage_ChargesFlattenedWorkDuringRectangleTextClipIntersections() {
        var commands = new List<OfficePathCommand> {
            OfficePathCommand.MoveTo(new OfficePoint(0D, 0D))
        };
        for (int index = 0; index < 11000; index++) {
            double x = index % 100;
            commands.Add(OfficePathCommand.CubicBezierTo(
                new OfficePoint(x + 0.25D, 25D),
                new OfficePoint(x + 0.75D, 75D),
                new OfficePoint(x + 1D, 100D)));
        }
        commands.Add(OfficePathCommand.Close());
        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath path));
        path = path.AsTextClippingPath();
        PdfPageClipPath rectangle = PdfPageClipPath.Rectangle(0D, 0D, 100D, 100D);
        var budget = new PdfTextClippingBudget();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            budget.ResolveActiveClip(path, rectangle));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPage_RejectsArbitraryPathFlatteningBeforeContourMaterialization() {
        var commands = new List<OfficePathCommand> {
            OfficePathCommand.MoveTo(new OfficePoint(0D, 0D))
        };
        for (int index = 0; index < 42000; index++) {
            double x = index % 100;
            commands.Add(OfficePathCommand.QuadraticBezierTo(
                new OfficePoint(x + 0.5D, 50D),
                new OfficePoint(x + 1D, 100D)));
        }
        commands.Add(OfficePathCommand.Close());
        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath active));
        active = active.AsTextClippingPath();
        Assert.True(PdfPageClipPath.TryCreatePath(new[] {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(100D, 0D),
            OfficePathCommand.LineTo(100D, 100D),
            OfficePathCommand.Close()
        }, OfficeFillRule.NonZero, out PdfPageClipPath next));
        var budget = new PdfTextClippingBudget();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            budget.ResolveActiveClip(active, next));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPage_ChargesGrowingPolygonClipWork() {
        Assert.True(PdfPageClipPath.TryCreatePath(new[] {
            OfficePathCommand.MoveTo(-10000D, -10000D),
            OfficePathCommand.LineTo(10000D, -10000D),
            OfficePathCommand.LineTo(0D, 10000D),
            OfficePathCommand.Close()
        }, OfficeFillRule.NonZero, out PdfPageClipPath active));
        active = active.AsTextClippingPath();

        const int clipVertices = 1500;
        var clipCommands = new List<OfficePathCommand>(clipVertices + 2);
        for (int index = 0; index < clipVertices; index++) {
            double angle = index * Math.PI * 2D / clipVertices;
            var point = new OfficePoint(Math.Cos(angle) * 100D, Math.Sin(angle) * 100D);
            clipCommands.Add(index == 0
                ? OfficePathCommand.MoveTo(point)
                : OfficePathCommand.LineTo(point));
        }
        clipCommands.Add(OfficePathCommand.Close());
        Assert.True(PdfPageClipPath.TryCreatePath(clipCommands, OfficeFillRule.NonZero, out PdfPageClipPath next));
        var budget = new PdfTextClippingBudget();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            budget.ResolveActiveClip(active, next));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void TextParser_ChargesLaterPathClipAgainstActiveTextClip() {
        string content = BuildTextClipFollowedByCurveHeavyPath();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                content,
                (_, bytes) => Encoding.ASCII.GetString(bytes),
                (_, bytes) => bytes.Length * 500D,
                pageHeight: 200D,
                textClippingBudget: new PdfTextClippingBudget()));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void InvocationParser_ChargesLaterPathClipAgainstActiveTextClip() {
        string content = BuildTextClipFollowedByCurveHeavyPath();
        var fonts = new Dictionary<string, PdfFontResource>(StringComparer.Ordinal) {
            ["F1"] = new PdfFontResource("F1", "Helvetica", "WinAnsiEncoding", hasToUnicode: false)
        };
        var widthProviders = new Dictionary<string, Func<byte[], double>>(StringComparer.Ordinal) {
            ["F1"] = bytes => bytes.Length * 500D
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageXObjectInvocationParser.Parse(
                content,
                Matrix2D.Identity,
                200D,
                graphicsStates: null,
                colorSpaces: null,
                fonts: fonts,
                fontWidthProviders: widthProviders,
                textClippingBudget: new PdfTextClippingBudget()));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void VisualParser_ChargesLaterPathClipAgainstInheritedTextClip() {
        string content = BuildCurveHeavyPathClip();
        PdfPageClipPath inheritedTextClip = PdfPageClipPath.Rectangle(0D, 100D, 100D, 120D).AsTextClippingPath();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageContentVisualParser.Parse(
                content,
                100D,
                200D,
                graphicsStates: null,
                colorSpaces: null,
                shadings: null,
                shadingPatterns: null,
                tilingPatterns: null,
                initialClipPath: inheritedTextClip,
                textClippingBudget: new PdfTextClippingBudget()));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void FormInvocationParser_ChargesLaterPathClipAgainstInheritedTextClip() {
        string content = BuildCurveHeavyPathClip();
        PdfPageClipPath inheritedTextClip = PdfPageClipPath.Rectangle(0D, 100D, 100D, 120D).AsTextClippingPath();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.ExtractFormInvocations(
                content,
                pageHeight: 200D,
                initialClipPath: inheritedTextClip,
                textClippingBudget: new PdfTextClippingBudget()));

        Assert.Equal(PdfReadLimitKind.TextClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumTextClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void ParserPassesBoundOrdinaryClipIntersectionWork() {
        string content = "0 -20 100 120 re W n " + BuildCurveHeavyPathClip();

        AssertIntersectionLimit(() => TextContentParser.Parse(
            content,
            (_, bytes) => Encoding.ASCII.GetString(bytes),
            (_, bytes) => bytes.Length * 500D,
            pageHeight: 200D,
            textClippingBudget: new PdfTextClippingBudget()));
        AssertIntersectionLimit(() => PdfPageXObjectInvocationParser.Parse(
            content,
            Matrix2D.Identity,
            200D,
            graphicsStates: null,
            colorSpaces: null,
            textClippingBudget: new PdfTextClippingBudget()));
        AssertIntersectionLimit(() => PdfPageContentVisualParser.Parse(
            content,
            100D,
            200D,
            graphicsStates: null,
            colorSpaces: null,
            shadings: null,
            shadingPatterns: null,
            tilingPatterns: null,
            textClippingBudget: new PdfTextClippingBudget()));
        AssertIntersectionLimit(() => TextContentParser.ExtractFormInvocations(
            content,
            pageHeight: 200D,
            textClippingBudget: new PdfTextClippingBudget()));

        static void AssertIntersectionLimit(Action action) {
            PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(action);
            Assert.Equal(PdfReadLimitKind.ClippingIntersectionWork, exception.Kind);
            Assert.Equal(PdfPageClipPath.MaximumClippingIntersectionWork, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
        }
    }

    [Fact]
    public void ParserChargesClippedContourRepresentabilityAgainstSourceEdges() {
        const int pointCount = 1100;
        var content = new StringBuilder();
        for (int index = 0; index < pointCount; index++) {
            double angle = index * Math.PI * 2D / pointCount;
            double x = 100D + (50D * Math.Cos(angle));
            double y = 100D + (50D * Math.Sin(angle));
            content.Append(x.ToString("0.######", CultureInfo.InvariantCulture))
                .Append(' ')
                .Append(y.ToString("0.######", CultureInfo.InvariantCulture))
                .Append(index == 0 ? " m " : " l ");
        }
        content.Append("h W n 0 0 200 200 re W n");

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                content.ToString(),
                (_, bytes) => Encoding.ASCII.GetString(bytes),
                (_, bytes) => bytes.Length * 500D,
                pageHeight: 200D,
                textClippingBudget: new PdfTextClippingBudget()));

        Assert.Equal(PdfReadLimitKind.ClippingIntersectionWork, exception.Kind);
        Assert.Equal(PdfPageClipPath.MaximumClippingIntersectionWork, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void TextParserAllowsBoundedOrdinaryClipAfterRestoringPreTextClipState() {
        const string content = "q BT /F1 12 Tf 4 Tr (A) Tj ET Q 0 -20 100 120 re W n 0 -20 100 120 re W n";

        _ = TextContentParser.Parse(
            content,
            (_, bytes) => Encoding.ASCII.GetString(bytes),
            (_, bytes) => bytes.Length * 500D,
            pageHeight: 200D,
            textClippingBudget: new PdfTextClippingBudget());
    }

    [Fact]
    public void RenderPage_DoesNotChargeDisjointTextClipIntersections() {
        const int runsPerObject = 1416;
        string runs = string.Concat(Enumerable.Repeat("(A) Tj ", runsPerObject));
        byte[] pdf = BuildSingleStreamPdf(
            "BT /F1 12 Tf 4 Tr " + runs + "ET BT /F1 12 Tf 1000 1000 Td 4 Tr " + runs + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        OfficeDrawing drawing = document.Pages[0].ToDrawing();

        Assert.NotNull(drawing);
    }

    private static string BuildTextClipFollowedByCurveHeavyPath() =>
        "BT /F1 12 Tf 4 Tr (A) Tj ET " + BuildCurveHeavyPathClip();

    private static string BuildCurveHeavyPathClip(int curveCount = 42000) {
        var content = new StringBuilder("0 -20 m ");
        for (int index = 0; index < curveCount; index++) {
            content.Append("0 -20 50 100 100 -20 c ");
        }
        return content.Append("h W n").ToString();
    }

    [Fact]
    public void ImagePlacementParserSharesTextClipBudgetAcrossRepeatedForms() {
        string textShows = string.Concat(Enumerable.Repeat("(A) Tj ", 1400));
        string form = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>",
            "BT /F1 12 Tf 4 Tr " + textShows + "ET");
        byte[] pdf = BuildSingleStreamPdf(
            "/Fm1 Do /Fm1 Do /Fm1 Do",
            "<< /XObject << /Fm1 5 0 R >> >>",
            form);
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].GetImagePlacements());

        Assert.True(
            exception.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected limit kind: {exception.Kind}.");
        Assert.True(exception.Actual > exception.Limit);

        PdfReadLimitException diagnosticException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).AssessRenderCompatibility());
        Assert.True(
            diagnosticException.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected diagnostic limit kind: {diagnosticException.Kind}.");
        Assert.True(diagnosticException.Actual > diagnosticException.Limit);
    }

    [Fact]
    public void RenderPageSharesTextClipBudgetAcrossType3GlyphPrograms() {
        string textShows = string.Concat(Enumerable.Repeat("(A) Tj ", 3000));
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Fm1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Fm1 Do");
        string form = BuildStreamObject(
            7,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>",
            "BT /F1 12 Tf 4 Tr " + textShows + "ET");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyph,
            form);
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.True(
            exception.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected limit kind: {exception.Kind}.");
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPageSharesTextClipBudgetAcrossTilingPatternSurfaces() {
        string textShows = string.Concat(Enumerable.Repeat("(A) Tj ", 3000));
        string pattern1 = BuildStreamObject(
            5,
            "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>",
            "BT /F1 1 Tf 4 Tr " + textShows + "ET");
        string pattern2 = BuildStreamObject(
            6,
            "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>",
            "BT /F1 1 Tf 4 Tr " + textShows + "ET");
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs /P1 scn 0 0 10 10 re f /P2 scn 10 0 10 10 re f",
            "<< /Pattern << /P1 5 0 R /P2 6 0 R >> >>",
            pattern1,
            pattern2);
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.True(
            exception.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected limit kind: {exception.Kind}.");
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPageSharesTextClipBudgetAcrossSoftMaskSurfaces() {
        string textShows = string.Concat(Enumerable.Repeat("(A) Tj ", 3000));
        string graphicsState1 = "5 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 7 0 R >> >>\nendobj";
        string graphicsState2 = "6 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string mask1 = BuildStreamObject(
            7,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency >> /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>",
            "BT /F1 1 Tf 4 Tr " + textShows + "ET");
        string mask2 = BuildStreamObject(
            8,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency >> /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>",
            "BT /F1 1 Tf 4 Tr " + textShows + "ET");
        byte[] pdf = BuildSingleStreamPdf(
            "/GS1 gs 0 0 10 10 re f /GS2 gs 10 0 10 10 re f",
            "<< /ExtGState << /GS1 5 0 R /GS2 6 0 R >> >>",
            graphicsState1,
            graphicsState2,
            mask1,
            mask2);
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.True(
            exception.Kind is PdfReadLimitKind.TextClippingPaths or PdfReadLimitKind.TextClippingIntersectionWork,
            $"Unexpected limit kind: {exception.Kind}.");
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPageSharesClipIntersectionBudgetWithType3LuminosityProofs() {
        string heavyClip = "0 -20 100 120 re W n " + BuildCurveHeavyPathClip(11000);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R /B 7 0 R >> /Encoding << /Differences [65 /A /B] >> /FirstChar 65 /LastChar 66 /Widths [500 500] /Resources << /ExtGState << /GS1 8 0 R /GS2 9 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string glyphB = BuildStreamObject(7, "<<", "500 0 d0 /GS2 gs 0 0 500 700 re f");
        string state1 = "8 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 10 0 R >> >>\nendobj";
        string state2 = "9 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 11 0 R >> >>\nendobj";
        string mask1 = BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /Type /Group /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", heavyClip + " 0 0 500 700 re f");
        string mask2 = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /Type /Group /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", heavyClip + " 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (AB) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, glyphB, state1, state2, mask1, mask2);
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.ClippingIntersectionWork, exception.Kind);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RenderPageDoesNotChargeSkippedOverlappingContourIntersections() {
        string activeClip = "0 -20 m " + string.Concat(Enumerable.Range(1, 300)
            .Select(index => index + " " + (index % 2 == 0 ? -20 : 100) + " l ")) + "h W n ";
        string coincidentShows = string.Concat(Enumerable.Repeat("(A) Tj 1 0 0 1 0 0 Tm ", 1000));
        byte[] pdf = BuildSingleStreamPdf(
            activeClip + "BT /F1 12 Tf 4 Tr " + coincidentShows + "ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        OfficeDrawing drawing = document.Pages[0].ToDrawing();

        Assert.NotNull(drawing);
    }


}
