using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextEditorTests {
    [Fact]
    public void InspectAndFindExposeBoundedTextGeometryAndStyle() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Alpha beta alphabet"))
            .ToBytes();
        IReadOnlyList<PdfTextSpan> spans = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans();
        PdfPageRegion region = RegionAround(spans);

        PdfRegionText inspected = PdfDocument.Open(pdf).Text.Inspect(region);
        IReadOnlyList<PdfTextMatch> contains = PdfDocument.Open(pdf).Text.Find("alpha");
        IReadOnlyList<PdfTextMatch> words = PdfDocument.Open(pdf).Text.Find("alpha", new PdfTextSearchOptions { WholeWords = true });

        Assert.Contains("Alpha beta alphabet", inspected.Text, StringComparison.Ordinal);
        Assert.NotEmpty(inspected.Spans);
        Assert.True(inspected.FontSize > 0D);
        Assert.Equal(2, contains.Count);
        PdfTextMatch wholeWord = Assert.Single(words);
        Assert.Equal("Alpha", wholeWord.Text, ignoreCase: true);
        Assert.True(wholeWord.Width > 0D);
        Assert.True(wholeWord.Height > 0D);
    }

    [Fact]
    public void InspectTreatsGenericSansSerifFontsAsHelvetica() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (sans source) Tj ET\n",
            firstBaseFont: "GenericSansSerif");

        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find(
            "sans source",
            new PdfTextSearchOptions { MatchCase = true }));

        Assert.Equal(PdfStandardFont.Helvetica, match.SuggestedFont);
    }

    [Fact]
    public void FindAndReplaceAllBoundMaterializedMatches() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (aaaa) Tj ET\n");
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxTextSearchMatches = 2 } };

        PdfReadLimitException findException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Open(source).Text.Find("a", new PdfTextSearchOptions { MatchCase = true }, options));
        PdfReadLimitException replaceException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Open(source).Text.ReplaceAll("a", "b", new PdfTextSearchOptions { MatchCase = true }, readOptions: options));

        Assert.Equal(PdfReadLimitKind.TextSearchMatches, findException.Kind);
        Assert.Equal(PdfReadLimitKind.TextSearchMatches, replaceException.Kind);
        Assert.Equal(2, findException.Limit);
    }

    [Fact]
    public void PortableTextRestampsRejectAuthoredRenderingIntent() {
        byte[] source = BuildRawTextPdf("q /Perceptual ri BT /F1 12 Tf 50 700 Td (managed color) Tj ET Q\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("managed color", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        NotSupportedException moveException = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Open(source).Text.Move(region, 10D, 0D));
        NotSupportedException replaceException = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Open(source).Text.Replace(region, "replacement"));

        Assert.Contains("cannot be recreated", moveException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("cannot be recreated", replaceException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplaceRemovesOnlyTextAndPreservesIntersectingAnnotation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Keep line above"))
            .Paragraph(paragraph => paragraph.Text("Replace this sentence"))
            .Paragraph(paragraph => paragraph.Text("Keep line below"))
            .ToBytes();
        PdfTextMatch sourceMatch = Assert.Single(PdfDocument.Open(source).Text.Find("Replace this sentence", new PdfTextSearchOptions { MatchCase = true }));
        PdfPageRegion region = new PdfPageRegion(1, sourceMatch.X, sourceMatch.Y, sourceMatch.Width, sourceMatch.Height);
        PdfAnnotationEditResult annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Highlight",
            Rectangle = new[] { region.X, region.Y, region.Right, region.Top },
            QuadPoints = new[] { region.X, region.Top, region.Right, region.Top, region.X, region.Y, region.Right, region.Y }
        });

        PdfTextEditResult result = PdfDocument.Open(annotated.Bytes).Text.Replace(region, "Replacement text");
        string text = result.Document.Read.Text();

        Assert.True(result.AffectedCount > 0);
        Assert.DoesNotContain("Replace this sentence", text, StringComparison.Ordinal);
        Assert.Contains("Replacement text", text, StringComparison.Ordinal);
        Assert.Contains("Keep line above", text, StringComparison.Ordinal);
        Assert.Contains("Keep line below", text, StringComparison.Ordinal);
        Assert.Single(result.Document.Read.AnnotationsBySubtype("Highlight"));
    }

    [Fact]
    public void MoveAndAddProduceInspectableTextAtRequestedLocations() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Stationary line above"))
            .Paragraph(paragraph => paragraph.Text("Move me"))
            .Paragraph(paragraph => paragraph.Text("Stationary line below"))
            .ToBytes();
        PdfTextMatch originalMatch = Assert.Single(PdfDocument.Open(source).Text.Find("Move me", new PdfTextSearchOptions { MatchCase = true }));
        PdfPageRegion sourceRegion = new PdfPageRegion(1, originalMatch.X, originalMatch.Y, originalMatch.Width, originalMatch.Height);

        PdfTextEditResult moved = PdfDocument.Open(source).Text.Move(sourceRegion, 80D, -40D);
        IReadOnlyList<PdfTextMatch> movedMatches = moved.Document.Text.Find("Move me", new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch movedMatch = Assert.Single(movedMatches);
        Assert.InRange(movedMatch.X, originalMatch.X + 79D, originalMatch.X + 81D);
        Assert.InRange(movedMatch.Y, sourceRegion.Y - 41D, sourceRegion.Top - 39D);
        Assert.Contains("Stationary line above", moved.Document.Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Stationary line below", moved.Document.Read.Text(), StringComparison.Ordinal);

        var addRegion = new PdfPageRegion(1, 72D, 120D, 200D, 40D);
        PdfTextEditResult added = moved.Document.Text.Add(addRegion, "Added caption", new PdfTextEditOptions {
            Font = PdfStandardFont.CourierBold,
            FontSize = 14D,
            Color = PdfColor.FromRgb(10, 80, 160)
        });
        PdfTextMatch addedMatch = Assert.Single(added.Document.Text.Find("Added caption", new PdfTextSearchOptions { MatchCase = true }));
        Assert.InRange(addedMatch.X, 71.9D, 72.1D);
        Assert.Equal(PdfStandardFont.CourierBold, addedMatch.SuggestedFont);
    }

    [Fact]
    public void ReplaceAllPreservesUnmatchedTextInTheSameSourceSpan() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Keep heading"))
            .Paragraph(paragraph => paragraph.Text("cat cat dog"))
            .Paragraph(paragraph => paragraph.Text("Keep footer"))
            .ToBytes();

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            "fox",
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true });
        string text = result.Document.Read.Text();

        Assert.Equal(2, result.AffectedCount);
        Assert.Contains("fox fox dog", text, StringComparison.Ordinal);
        Assert.DoesNotContain("cat", text, StringComparison.Ordinal);
        Assert.Contains("Keep heading", text, StringComparison.Ordinal);
        Assert.Contains("Keep footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceAllKeepsSameBaselineColumnsSeparateAndUsesMatchedSpanStyle() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (left cat) Tj ET\n" +
            "BT /F2 20 Tf 350 700 Td (right cat) Tj ET\n");

        IReadOnlyList<PdfTextMatch> matches = PdfDocument.Open(source).Text.Find("cat", new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch right = Assert.Single(matches, static match => match.X > 300D);
        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll("left cat", "left fox", new PdfTextSearchOptions { MatchCase = true });
        string text = result.Document.Read.Text();

        Assert.Equal(PdfStandardFont.Courier, right.SuggestedFont);
        Assert.Equal(20D, right.FontSize, 2);
        Assert.Contains("left fox", text, StringComparison.Ordinal);
        Assert.Contains("right cat", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceAllDoesNotReflowIndependentColumns() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (cat) Tj ET\n" +
            "BT /F1 12 Tf 350 700 Td (cat) Tj ET\n");
        PdfTextMatch originalRight = Assert.Single(
            PdfDocument.Open(source).Text.Find("cat", new PdfTextSearchOptions { MatchCase = true }),
            static match => match.X > 300D);

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            "a much longer replacement",
            new PdfTextSearchOptions { MatchCase = true });
        IReadOnlyList<PdfTextMatch> replacements = result.Document.Text.Find(
            "a much longer replacement",
            new PdfTextSearchOptions { MatchCase = true });

        Assert.Equal(2, result.AffectedCount);
        Assert.Equal(2, replacements.Count);
        PdfTextMatch rewrittenRight = Assert.Single(replacements, static match => match.X > 300D);
        Assert.InRange(rewrittenRight.X, originalRight.X - 0.1D, originalRight.X + 0.1D);
    }

    [Fact]
    public void ReplaceAllPreservesExactUnmatchedWhitespaceAndCarriesInputBudgetAcrossRewrites() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (cat  cat dog) Tj ET\n");
        var readOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length } };
        string decodedSource = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans()).Text;

        PdfTextEditResult result = PdfDocument.Open(source, readOptions).Text.ReplaceAll(
            "cat",
            "longer-fox",
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true });
        string rewritten = string.Concat(PdfReadDocument.Open(result.Document.ToBytes()).Pages[0]
            .GetTextSpans()
            .OrderBy(static span => span.X)
            .Select(static span => span.Text));

        Assert.Equal(decodedSource.Replace("cat", "longer-fox"), rewritten);
        Assert.Equal(2, result.AffectedCount);
    }

    [Fact]
    public void SearchExcludesInvisibleAndClippedTextAndMutationFailsClosedWhenAtomicRemovalWouldExposeIt() {
        byte[] invisible = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (visible) Tj 3 Tr ( hidden secret) Tj 0 Tr ET\n");
        byte[] clipped = BuildRawTextPdf(
            "q 0 0 10 10 re W n BT /F1 12 Tf 50 700 Td (clipped secret) Tj ET Q\n");

        Assert.Empty(PdfDocument.Open(invisible).Text.Find("secret", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Empty(PdfDocument.Open(clipped).Text.Find("secret", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch visible = Assert.Single(PdfDocument.Open(invisible).Text.Find("visible", new PdfTextSearchOptions { MatchCase = true }));
        var visibleRegion = new PdfPageRegion(1, visible.X, visible.Y, visible.Width, visible.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(invisible).Text.Replace(visibleRegion, "updated"));
    }

    [Fact]
    public void RotatedMatchUsesMatchedSliceGeometryAndPreservesRotationDuringReplacement() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 0 1 -1 0 200 300 Tm (rotate cat) Tj ET\n");

        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("cat", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll("cat", "fox", new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch replacement = Assert.Single(result.Document.Text.Find("fox", new PdfTextSearchOptions { MatchCase = true }));

        Assert.InRange(match.RotationDegrees, 89.9D, 90.1D);
        Assert.True(match.Height > match.Width);
        Assert.InRange(replacement.RotationDegrees, 89.9D, 90.1D);
    }

    [Fact]
    public void ReplaceAllMapsCrossSpanPhraseEditsWithoutNormalizingUnmatchedSpanText() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (alpha) Tj 38 0 Td (beta tail) Tj ET\n");

        PdfTextMatch phrase = Assert.Single(PdfDocument.Open(source).Text.Find("alpha beta", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll("alpha beta", "gamma", new PdfTextSearchOptions { MatchCase = true });
        string text = result.Document.Read.Text();

        Assert.True(phrase.Width > 38D);
        Assert.Contains("gamma", text, StringComparison.Ordinal);
        Assert.Contains("tail", text, StringComparison.Ordinal);
        Assert.DoesNotContain("alpha", text, StringComparison.Ordinal);
        Assert.DoesNotContain("beta", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FindUsesDecodedGlyphAdvancesForSubstringGeometry() {
        byte[] source = BuildRawTextPdf("BT /F1 20 Tf 50 700 Td (iiiiWWWW) Tj ET\n");

        PdfTextMatch narrow = Assert.Single(PdfDocument.Open(source).Text.Find("iiii", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch wide = Assert.Single(PdfDocument.Open(source).Text.Find("WWWW", new PdfTextSearchOptions { MatchCase = true }));

        Assert.True(wide.Width > narrow.Width * 2D);
    }

    [Fact]
    public void FindGroupsRotatedSpansByTheirProjectedBaseline() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 0 1 -1 0 200 300 Tm (alpha) Tj ET\n" +
            "BT /F1 12 Tf 0 1 -1 0 200 340 Tm (beta) Tj ET\n");

        PdfTextMatch phrase = Assert.Single(PdfDocument.Open(source).Text.Find("alpha beta", new PdfTextSearchOptions { MatchCase = true }));

        Assert.InRange(phrase.RotationDegrees, 89.9D, 90.1D);
    }

    [Fact]
    public void InspectOrdersRotatedSpansAlongTheirProjectedBaseline() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 0 1 -1 0 200 300 Tm (alpha) Tj ET\n" +
            "BT /F1 12 Tf 0 1 -1 0 200 340 Tm (beta) Tj ET\n");

        PdfRegionText inspected = PdfDocument.Open(source).Text.Inspect(
            new PdfPageRegion(1, 180D, 280D, 50D, 100D));

        Assert.Equal("alpha beta", inspected.Text);
    }

    [Fact]
    public void MovePreservesIndependentSpanStyleAndPlacement() {
        byte[] source = BuildRawTextPdf(
            "1 0 0 rg BT /F1 12 Tf 50 700 Td (red) Tj ET\n" +
            "0 0 1 rg BT /F2 18 Tf 90 700 Td (blue) Tj ET\n");
        var region = new PdfPageRegion(1, 45D, 680D, 140D, 45D);

        PdfTextEditResult result = PdfDocument.Open(source).Text.Move(region, 40D, -30D);
        PdfReadPage page = PdfReadDocument.Open(result.Document.ToBytes()).Pages[0];
        PdfTextSpan red = Assert.Single(page.GetTextSpans(), static span => span.Text == "red");
        PdfTextSpan blue = Assert.Single(page.GetTextSpans(), static span => span.Text == "blue");

        Assert.InRange(red.X, 89.9D, 90.1D);
        Assert.InRange(blue.X, 129.9D, 130.1D);
        Assert.Equal(12D, red.FontSize, 2);
        Assert.Equal(18D, blue.FontSize, 2);
        Assert.True(red.Color!.Value.R > red.Color.Value.B);
        Assert.True(blue.Color!.Value.B > blue.Color.Value.R);
    }

    [Fact]
    public void InspectDominantStyleIncludesPaintColor() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 0 0 1 rg 50 700 Td (x) Tj ET\n" +
            "BT /F1 12 Tf 1 0 0 rg 65 700 Td (dominant-red) Tj ET\n");

        PdfRegionText inspected = PdfDocument.Open(source).Text.Inspect(new PdfPageRegion(1, 45D, 680D, 180D, 45D));

        Assert.True(inspected.Color.R > inspected.Color.B);
    }

    [Fact]
    public void ReplaceAllReflowsTrailingCrossSpanSuffixAndUsesOneBatchStamp() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (alpha) Tj 38 0 Td (beta tail) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll("alpha beta", "a much wider replacement", new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch replacement = Assert.Single(result.Document.Text.Find("a much wider replacement", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch tail = Assert.Single(result.Document.Text.Find("tail", new PdfTextSearchOptions { MatchCase = true }));
        string syntax = System.Text.Encoding.GetEncoding(28591).GetString(result.Document.ToBytes());

        Assert.True(tail.X >= replacement.X + replacement.Width - 1D);
        Assert.Contains("/OIMOEditF1", syntax, StringComparison.Ordinal);
        Assert.DoesNotContain("/OIMOEditF2", syntax, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceAllReflowsAnUnmatchedTrailingSpanInTheSameTextFlow() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (cat) Tj ET\n" +
            "BT /F1 12 Tf 75 700 Td (tail) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            "a much wider replacement",
            new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch replacement = Assert.Single(result.Document.Text.Find(
            "a much wider replacement",
            new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch tail = Assert.Single(result.Document.Text.Find(
            "tail",
            new PdfTextSearchOptions { MatchCase = true }));

        Assert.True(tail.X >= replacement.X + replacement.Width - 1D);
    }

    [Theory]
    [InlineData("1 Tr")]
    [InlineData("2 Tr")]
    [InlineData("4 Tr")]
    [InlineData("5 Tr")]
    [InlineData("6 Tr")]
    public void MutationRejectsTextRenderingModesTheFillOnlyStamperCannotPreserve(string renderingMode) {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf " + renderingMode + " 50 700 Td (clip source) Tj ET\n");
        var region = new PdfPageRegion(1, 45D, 680D, 120D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "replacement"));
    }

    [Theory]
    [InlineData("2 0 0 1 50 700 Tm")]
    [InlineData("-1 0 0 1 150 700 Tm")]
    [InlineData("1 .3 0 1 50 700 Tm")]
    public void MutationRejectsTextTransformsTheStandardFontStamperCannotPreserve(string matrix) {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf " + matrix + " (transformed) Tj ET\n");
        var region = new PdfPageRegion(1, 0D, 650D, 300D, 100D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "replacement"));
    }

    [Fact]
    public void ConformalTextScaleIsPreservedAndMutationRecordsContentOperationAndReadOptions() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 2 0 0 2 50 350 Tm (scaled) Tj ET\n");
        var readOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxContentOperations = 321 } };
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("scaled", new PdfTextSearchOptions { MatchCase = true }));

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll("scaled", "updated", readOptions: readOptions);
        PdfTextMatch updated = Assert.Single(result.Document.Text.Find("updated", new PdfTextSearchOptions { MatchCase = true }));
        PdfPipelineStep mutation = Assert.Single(result.Document.Pipeline.Steps, static step => step.Kind == PdfPipelineStepKind.Mutation);

        Assert.Equal(24D, match.FontSize, 2);
        Assert.Equal(24D, updated.FontSize, 2);
        Assert.Equal(PdfMutationOperation.ModifyPageContent, mutation.MutationOperation);
        Assert.Equal(321, result.Document.ReadOptions.Limits.MaxContentOperations);
    }

    [Fact]
    public void MutationFailsClosedWhenAppendingWouldReverseOverlappingPaintOrder() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 20 Tf 50 700 Td (covered) Tj ET\n" +
            "1 1 1 rg 45 685 100 30 re f\n");
        var region = new PdfPageRegion(1, 45D, 680D, 110D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "changed"));
    }

    [Fact]
    public void MoveChecksTheProjectedDestinationAgainstLaterPaint() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 20 Tf 50 700 Td (move me) Tj ET\n" +
            "1 1 1 rg 295 685 110 30 re f\n");
        var region = new PdfPageRegion(1, 45D, 680D, 120D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Move(region, 250D, 0D));
    }

    [Fact]
    public void RotatedPageMoveChecksDestinationInVisualCoordinates() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 20 Tf 50 700 Td (move me) Tj ET\n" +
            "1 1 1 rg 295 685 110 30 re f\n",
            pageEntries: "/Rotate 90");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("move me", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X - 5D, match.Y - 5D, match.Width + 10D, match.Height + 10D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Move(region, 250D, 0D));
    }

    [Fact]
    public void ConformalScaleUsesTheRestampFontSizeForStackingBounds() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 2 0 0 2 50 350 Tm (scaled) Tj ET\n" +
            "1 1 1 rg 45 362 120 4 re f\n");
        var region = new PdfPageRegion(1, 45D, 325D, 150D, 60D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Theory]
    [InlineData("2 Tc (spaced) Tj")]
    [InlineData("4 Tw (spaced text) Tj")]
    [InlineData("[(spa) -120 (ced)] TJ")]
    public void MutationRejectsTextSpacingThePlainStamperCannotRecreate(string showText) {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td " + showText + " ET\n");
        var region = new PdfPageRegion(1, 45D, 680D, 180D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void FindReturnsVisibleTextThatMutationCannotSafelyRestamp() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td 2 Tc (spaced) Tj ET\n");

        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find(
            "spaced",
            new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void WholeWordSearchDoesNotSplitADecomposedGrapheme() {
        const string decomposed = "re\u0301sume\u0301";
        byte[] source = BuildRawTextPdf(
            "/Span << /ActualText <FEFF00720065030100730075006D00650301> >> BDC " +
            "BT /F1 12 Tf 50 700 Td (resume) Tj ET EMC\n");

        Assert.Empty(PdfDocument.Open(source).Text.Find(
            "re",
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true }));
        Assert.Single(PdfDocument.Open(source).Text.Find(
            decomposed,
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true }));
    }

    [Fact]
    public void WholeWordSearchDoesNotSplitBeforeASupplementaryPlaneLetter() {
        const string text = "cat\U00010400";
        byte[] source = BuildRawTextPdf(
            "/Span << /ActualText <FEFF006300610074D801DC00> >> BDC " +
            "BT /F1 12 Tf 50 700 Td (catX) Tj ET EMC\n");

        Assert.Empty(PdfDocument.Open(source).Text.Find(
            "cat",
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true }));
        Assert.Single(PdfDocument.Open(source).Text.Find(
            text,
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true }));
    }

    [Fact]
    public void SearchExcludesInferredSpacesAtMatchBoundaries() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (alpha) Tj 38 0 Td (beta) Tj ET\n");

        Assert.Empty(PdfDocument.Open(source).Text.Find(" ", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Empty(PdfDocument.Open(source).Text.Find(" beta", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Single(PdfDocument.Open(source).Text.Find("alpha beta", new PdfTextSearchOptions { MatchCase = true }));
    }

    [Fact]
    public void ReplaceAllAppliesOverridesOnlyToReplacementFragments() {
        byte[] source = BuildRawTextPdf(
            "1 0 0 rg BT /F1 12 Tf 50 700 Td (cat tail) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            "a much wider replacement",
            new PdfTextSearchOptions { MatchCase = true },
            new PdfTextEditOptions { Color = PdfColor.FromRgb(0, 0, 255) });
        PdfTextMatch replacement = Assert.Single(result.Document.Text.Find("a much wider replacement", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch tail = Assert.Single(result.Document.Text.Find("tail", new PdfTextSearchOptions { MatchCase = true }));

        Assert.True(replacement.Color.B > replacement.Color.R);
        Assert.True(tail.Color.R > tail.Color.B);
        Assert.True(tail.X >= replacement.X + replacement.Width - 1D);
    }

    [Fact]
    public void SearchAndReflowUseAFontRelativeIndependentFlowCutoff() {
        byte[] source = BuildRawTextPdf(
            "BT /F2 72 Tf 50 700 Td (alpha) Tj ET\n" +
            "BT /F2 72 Tf 310 700 Td (beta) Tj ET\n");

        Assert.Single(PdfDocument.Open(source).Text.Find("alpha beta", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "alpha",
            "a",
            new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch beta = Assert.Single(result.Document.Text.Find("beta", new PdfTextSearchOptions { MatchCase = true }));

        Assert.True(beta.X < 310D);
    }

    [Fact]
    public void ReplaceSelectsATextSpanThatOnlyIntersectsTheRegionEdge() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (long source text) Tj ET\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find(
            "long source text",
            new PdfTextSearchOptions { MatchCase = true }));
        var edge = new PdfPageRegion(1, match.X + 0.25D, match.Y, 1D, match.Height);

        PdfRegionText inspected = PdfDocument.Open(source).Text.Inspect(edge);
        PdfTextEditResult result = PdfDocument.Open(source).Text.Replace(edge, "updated");

        Assert.Contains("long source text", inspected.Text, StringComparison.Ordinal);
        Assert.Equal(1, result.AffectedCount);
        Assert.DoesNotContain("long source text", result.Document.Read.Text(), StringComparison.Ordinal);
        Assert.Contains("updated", result.Document.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void MutationRejectsNonDefaultBlendAndSoftMaskTextEffects() {
        byte[] blend = BuildRawTextPdf(
            "/GSBlend gs BT /F1 12 Tf 50 700 Td (blended) Tj ET\n",
            "/ExtGState << /GSBlend << /BM /Multiply >> >>");
        byte[] masked = BuildRawTextPdf(
            "/GSMask gs BT /F1 12 Tf 50 700 Td (masked) Tj ET\n",
            "/ExtGState << /GSMask << /SMask << /S /Alpha /G 7 0 R >> >> >>",
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 600 800] /Group << /S /Transparency /CS /DeviceRGB >> /Length 0 >>\nstream\nendstream\nendobj\n");
        var region = new PdfPageRegion(1, 45D, 680D, 160D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(blend).Text.Replace(region, "updated"));
        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(masked).Text.Replace(region, "updated"));
    }

    [Fact]
    public void MutationRejectsUnsupportedBlendAndMalformedSoftMaskEffects() {
        byte[] blend = BuildRawTextPdf(
            "/GSBlend gs BT /F1 12 Tf 50 700 Td (blended) Tj ET\n",
            "/ExtGState << /GSBlend << /BM /UnsupportedMode >> >>");
        byte[] masked = BuildRawTextPdf(
            "/GSMask gs BT /F1 12 Tf 50 700 Td (masked) Tj ET\n",
            "/ExtGState << /GSMask << /SMask << /S /Alpha >> >> >>");
        var region = new PdfPageRegion(1, 45D, 680D, 160D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(blend).Text.Replace(region, "updated"));
        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(masked).Text.Replace(region, "updated"));
    }

    [Fact]
    public void MutationRejectsTextInAFormWithInheritedBlendEffect() {
        const string formText = "BT /F1 12 Tf 50 700 Td (form text) Tj ET\n";
        byte[] source = BuildRawTextPdf(
            "/GSBlend gs /Fm Do\n",
            "/ExtGState << /GSBlend << /BM /Multiply >> >> /XObject << /Fm 7 0 R >>",
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 600 800] /Resources << /Font << /F1 5 0 R >> >> /Length " + System.Text.Encoding.ASCII.GetByteCount(formText) + " >>\nstream\n" + formText + "endstream\nendobj\n");
        var region = new PdfPageRegion(1, 45D, 680D, 160D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void InnerGraphicsStateDoesNotClearAnInheritedUnsupportedFormEffect() {
        const string formText = "/GSOpacity gs BT /F1 12 Tf 50 700 Td (form text) Tj ET\n";
        byte[] source = BuildRawTextPdf(
            "/GSBlend gs /Fm Do\n",
            "/ExtGState << /GSBlend << /BM /Multiply >> >> /XObject << /Fm 7 0 R >>",
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 600 800] /Resources << /Font << /F1 5 0 R >> /ExtGState << /GSOpacity << /ca 0.5 >> >> >> /Length " + System.Text.Encoding.ASCII.GetByteCount(formText) + " >>\nstream\n" + formText + "endstream\nendobj\n");
        var region = new PdfPageRegion(1, 45D, 680D, 160D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void EditingEarlierTextDoesNotDuplicateASurvivingLaterSpan() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (edit) Tj ET\n" +
            "BT /F1 12 Tf 50 650 Td (survivor) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "edit",
            "updated",
            new PdfTextSearchOptions { MatchCase = true });

        Assert.Single(result.Document.Text.Find("survivor", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Single(PdfReadDocument.Open(result.Document.ToBytes()).Pages[0].GetTextSpans(), static span => span.Text == "survivor");
    }

    [Fact]
    public void MutationFailsClosedWhenAReferencedLaterPaintCannotBeBounded() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (source) Tj ET\n/Sh1 sh\n",
            "/Shading << /Sh1 7 0 R >>",
            "7 0 obj\n<< /ShadingType 4 /ColorSpace /DeviceRGB >>\nendobj\n");
        var region = new PdfPageRegion(1, 45D, 680D, 120D, 45D);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void MutationRejectsTextWhosePatternPaintCannotBeRestamped() {
        byte[] source = BuildRawTextPdf(
            "/PCS cs /P1 scn BT /F1 12 Tf 50 700 Td (pattern text) Tj ET\n",
            "/ColorSpace << /PCS /Pattern >> /Pattern << /P1 7 0 R >>",
            "7 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 8 8] /XStep 8 /YStep 8 /Resources << >> /Length 0 >>\nstream\nendstream\nendobj\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("pattern text", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void MutationRejectsTextAfterAnUnresolvedFillColorSpaceSelection() {
        byte[] source = BuildRawTextPdf(
            "1 0 0 rg /Missing cs 0 1 0 scn BT /F1 12 Tf 50 700 Td (unknown color) Tj ET\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("unknown color", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Theory]
    [InlineData("/Span << /MCID 0 >> BDC", "EMC")]
    [InlineData("/OC /Layer BDC", "EMC")]
    public void MutationRejectsTextBoundToTaggedOrOptionalContent(string beginMarkedContent, string endMarkedContent) {
        byte[] source = BuildRawTextPdf(
            beginMarkedContent + " BT /F1 12 Tf 50 700 Td (structured text) Tj ET " + endMarkedContent + "\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("structured text", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void MutationRejectsTextObjectsThatEstablishPersistentGraphicsState() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 1 0 0 rg 50 700 Td (stateful text) Tj ET\n" +
            "0 0 40 40 re f\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("stateful text", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated"));
    }

    [Fact]
    public void SearchAndInspectOrderRightToLeftSpansAlongDecreasingBaselines() {
        byte[] source = BuildRawTextPdf(
            "/Span << /ActualText <FEFF05E905DC05D505DD> >> BDC BT /F1 12 Tf 220 700 Td (x) Tj ET EMC\n" +
            "/Span << /ActualText <FEFF05E205D505DC05DD> >> BDC BT /F1 12 Tf 205 700 Td (x) Tj ET EMC\n");
        Assert.Equal(new[] { "שלום", "עולם" }, PdfReadDocument.Open(source).Pages[0].GetTextSpans().Select(static span => span.Text).ToArray());

        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("שלום עולם", new PdfTextSearchOptions { MatchCase = true }));
        PdfRegionText inspected = PdfDocument.Open(source).Text.Inspect(new PdfPageRegion(1, 190, 680, 60, 40));

        Assert.Equal("שלום עולם", match.Text);
        Assert.Equal("שלום עולם", inspected.Text);
    }

    [Fact]
    public void ReplaceAllContinuesSuffixFromTheFinalReplacementLine() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (cat tail) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll("cat", "foo\nbar", new PdfTextSearchOptions { MatchCase = true });

        Assert.Single(result.Document.Text.Find("bar tail", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch bar = Assert.Single(result.Document.Text.Find("bar", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch tail = Assert.Single(result.Document.Text.Find("tail", new PdfTextSearchOptions { MatchCase = true }));
        Assert.InRange(tail.Y, bar.Y - 0.5D, bar.Y + 0.5D);
        Assert.True(tail.X > bar.X);
    }

    [Theory]
    [InlineData("foo\n")]
    [InlineData("\n")]
    public void ReplaceAllContinuesSuffixAfterTrailingEmptyReplacementLine(string replacement) {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (cat tail) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            replacement,
            new PdfTextSearchOptions { MatchCase = true });

        PdfTextMatch tail = Assert.Single(result.Document.Text.Find("tail", new PdfTextSearchOptions { MatchCase = true }));
        Assert.True(tail.Y < 699D);
        Assert.InRange(tail.X, 49D, 53D);
    }

    [Fact]
    public void ReplaceAllProjectsRotatedReplacementAdvanceOntoSourceFlow() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (cat tail) Tj ET\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            "fox",
            new PdfTextSearchOptions { MatchCase = true },
            new PdfTextEditOptions { RotationDegrees = 90D });
        PdfTextMatch replacement = Assert.Single(result.Document.Text.Find("fox", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch tail = Assert.Single(result.Document.Text.Find("tail", new PdfTextSearchOptions { MatchCase = true }));

        Assert.InRange(replacement.RotationDegrees, 89.9D, 90.1D);
        Assert.InRange(tail.X, 49D, 53D);
        Assert.True(tail.Y > replacement.Y);
    }

    [Fact]
    public void SearchIncludesVisibleArtifactTextButMutationFailsClosed() {
        byte[] source = BuildRawTextPdf("/Artifact BMC BT /F1 12 Tf 50 700 Td (visible footer) Tj ET EMC\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("visible footer", new PdfTextSearchOptions { MatchCase = true }));
        var region = new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height);

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(region, "updated footer"));
    }

    [Fact]
    public void TextBatchStampIsolatesInheritedTransformAndClipState() {
        byte[] source = BuildRawTextPdf("2 0 0 2 0 0 cm 0 0 10 10 re W n\n");

        PdfTextEditResult result = PdfDocument.Open(source).Text.Add(
            new PdfPageRegion(1, 100D, 100D, 160D, 30D),
            "isolated stamp");
        PdfTextMatch match = Assert.Single(result.Document.Text.Find("isolated stamp", new PdfTextSearchOptions { MatchCase = true }));

        Assert.InRange(match.X, 99.9D, 100.1D);
        Assert.InRange(match.Y, 113D, 117D);
    }

    [Fact]
    public void MutationRejectsTextInheritedFromTransparencyGroupForm() {
        const string formText = "BT /F1 12 Tf 50 700 Td (group text) Tj ET\n";
        byte[] source = BuildRawTextPdf(
            "/Fm Do\n",
            "/XObject << /Fm 7 0 R >>",
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 600 800] /Group << /S /Transparency >> /Resources << /Font << /F1 5 0 R >> >> /Length " + System.Text.Encoding.ASCII.GetByteCount(formText) + " >>\nstream\n" + formText + "endstream\nendobj\n");
        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(
            new PdfPageRegion(1, 45, 680, 100, 40),
            "updated"));
    }

    [Fact]
    public void TextEditorCoordinatesAreRelativeToNonzeroPageBoxOrigin() {
        byte[] source = BuildRawTextPdf(
            "BT /F1 12 Tf 150 650 Td (existing) Tj ET\n",
            pageEntries: "/CropBox [100 100 500 700]");

        PdfTextMatch existing = Assert.Single(PdfDocument.Open(source).Text.Find("existing", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextEditResult replaced = PdfDocument.Open(source).Text.Replace(
            new PdfPageRegion(1, existing.X, existing.Y, existing.Width, existing.Height),
            "changed");
        PdfTextEditResult added = PdfDocument.Open(source).Text.Add(new PdfPageRegion(1, 0, 0, 120, 40), "origin text");
        PdfTextMatch addedMatch = Assert.Single(added.Document.Text.Find("origin text", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextMatch changed = Assert.Single(replaced.Document.Text.Find("changed", new PdfTextSearchOptions { MatchCase = true }));

        Assert.InRange(existing.X, 49.9D, 50.1D);
        Assert.InRange(existing.Y, 546D, 548D);
        Assert.InRange(addedMatch.X, -0.1D, 0.1D);
        Assert.InRange(addedMatch.Y, 24D, 26D);
        Assert.InRange(changed.X, existing.X - 0.1D, existing.X + 0.1D);
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r\n\n")]
    public void AddRejectsLineBreakOnlyText(string text) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Page")).ToBytes();

        Assert.Throws<ArgumentException>(() => PdfDocument.Open(source).Text.Add(new PdfPageRegion(1, 20, 20, 100, 30), text));
    }

    [Fact]
    public void MutationRejectsUnmodeledNonstrokingOverprintState() {
        byte[] source = BuildRawTextPdf(
            "/GSOverprint gs BT /F1 12 Tf 50 700 Td (overprint) Tj ET\n",
            "/ExtGState << /GSOverprint << /op true /OPM 1 >> >>");

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Replace(
            new PdfPageRegion(1, 45D, 680D, 160D, 45D),
            "updated"));
    }

    [Fact]
    public void MutationRejectsActualTextThatDiffersFromPaintedGlyphs() {
        byte[] source = BuildRawTextPdf(
            "/Span << /ActualText <FEFF006600660069> >> BDC " +
            "BT /F1 12 Tf 50 700 Td (x) Tj ET EMC\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("ffi", new PdfTextSearchOptions { MatchCase = true }));

        Assert.Throws<NotSupportedException>(() => PdfDocument.Open(source).Text.Move(
            new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height),
            20D,
            0D));
    }

    [Fact]
    public void MovePreservesAuthoredEdgeWhitespace() {
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td ( tail ) Tj ET\n");
        PdfTextMatch match = Assert.Single(PdfDocument.Open(source).Text.Find("tail", new PdfTextSearchOptions { MatchCase = true }));

        PdfTextEditResult result = PdfDocument.Open(source).Text.Move(
            new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height),
            20D,
            0D);

        Assert.Contains("<207461696C20> Tj", PdfEncoding.Latin1GetString(result.Document.ToBytes()), StringComparison.Ordinal);
        PdfTextSpan moved = Assert.Single(PdfReadDocument.Open(result.Document.ToBytes()).Pages[0].GetTextSpans(), static span => span.Text == "tail");
        Assert.Equal(" tail ", moved.RestampText);
    }

    [Fact]
    public void TextMutationsPreserveAnExistingAcroFormCatalogGraph() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Form and text edit")).ToBytes();
        byte[] form = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "customer.notes",
            Kind = PdfFormFieldCreationKind.Text,
            X = 72D,
            Y = 600D,
            Width = 180D,
            Height = 24D,
            Value = "kept"
        })).ToBytes();

        PdfTextEditResult added = PdfDocument.Open(form).Text.Add(new PdfPageRegion(1, 72D, 520D, 180D, 30D), "added text");
        PdfTextMatch sourceMatch = Assert.Single(added.Document.Text.Find("Form and text edit", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextEditResult moved = added.Document.Text.Move(
            new PdfPageRegion(1, sourceMatch.X, sourceMatch.Y, sourceMatch.Width, sourceMatch.Height),
            20D,
            -20D);
        PdfTextMatch movedMatch = Assert.Single(moved.Document.Text.Find("Form and text edit", new PdfTextSearchOptions { MatchCase = true }));
        PdfTextEditResult result = moved.Document.Text.Replace(
            new PdfPageRegion(1, movedMatch.X, movedMatch.Y, movedMatch.Width, movedMatch.Height),
            "updated text");

        PdfFormField field = Assert.Single(result.Document.Inspect().FormFields);
        Assert.Equal("customer.notes", field.Name);
        Assert.Equal("kept", field.Value);
        Assert.Single(result.Document.Text.Find("added text", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Single(result.Document.Text.Find("updated text", new PdfTextSearchOptions { MatchCase = true }));
    }

    [Fact]
    public void ReplaceCarriesConfiguredContentNestingLimitThroughTextRemoval() {
        string nestedOperand = new string('[', 129) + "0" + new string(']', 129) + " n\n";
        byte[] source = BuildRawTextPdf(nestedOperand + "BT /F1 12 Tf 50 700 Td (replace me) Tj ET\n");
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 130 }
        };
        PdfDocument document = PdfDocument.Open(source, readOptions);
        PdfTextMatch match = Assert.Single(document.Text.Find("replace me", new PdfTextSearchOptions { MatchCase = true }));

        PdfTextEditResult result = document.Text.Replace(
            new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height),
            "replaced");

        Assert.Single(result.Document.Text.Find("replaced", new PdfTextSearchOptions { MatchCase = true }));
    }

    [Fact]
    public void ReplaceUsesTokenizedTextObjectsAcrossNamesAndInlineImagePayloads() {
        byte[] source = BuildRawTextPdf(
            "BI /W 1 /H 1 /BPC 8 /CS /RGB ID BT EI " +
            "BT /F1 12 Tf 50 700 Td /ET MP (replace me) Tj ET\n");
        PdfDocument document = PdfDocument.Open(source);
        PdfTextMatch match = Assert.Single(document.Text.Find("replace me", new PdfTextSearchOptions { MatchCase = true }));

        PdfTextEditResult result = document.Text.Replace(
            new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height),
            "replaced");

        Assert.Empty(result.Document.Text.Find("replace me", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Single(result.Document.Text.Find("replaced", new PdfTextSearchOptions { MatchCase = true }));
    }

    [Fact]
    public void TextAddRejectsTaggedCatalogWithoutOwnedStructureAssociation() {
        byte[] raw = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (tagged source) Tj ET\n",
            additionalObjects: "7 0 obj\n<< /Type /StructTreeRoot /K [] >>\nendobj\n");
        string taggedText = PdfEncoding.Latin1GetString(raw).Replace(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 7 0 R >>");

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Open(PdfEncoding.Latin1GetBytes(taggedText)).Text.Add(
                new PdfPageRegion(1, 100D, 100D, 160D, 30D),
                "new text"));

        Assert.Contains("FullRewrite.TaggedContent", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void SearchOptionsSnapshotPagesAndRejectInvalidSelection() {
        int[] pages = { 1 };
        var options = new PdfTextSearchOptions { PageNumbers = pages };
        pages[0] = 2;

        Assert.Equal(1, Assert.Single(options.PageNumbers!));
        Assert.Throws<ArgumentOutOfRangeException>(() => options.PageNumbers = new[] { 0 });
        Assert.Throws<ArgumentException>(() => options.PageNumbers = new[] { 1, 1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageRegion(0, 0D, 0D, 10D, 10D));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextEditOptions { FontSize = 0D });
    }

    [Fact]
    public void PublicRedactionStillPaintsAndRemovesIntersectingAnnotations() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Secret marker"))
            .ToBytes();
        IReadOnlyList<PdfTextSpan> spans = PdfReadDocument.Open(source).Pages[0].GetTextSpans();
        PdfPageRegion region = RegionAround(spans);
        byte[] annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Highlight",
            Rectangle = new[] { region.X, region.Y, region.Right, region.Top },
            QuadPoints = new[] { region.X, region.Top, region.Right, region.Top, region.X, region.Y, region.Right, region.Y }
        }).Bytes;

        PdfDocument redacted = PdfDocument.Open(annotated).Redactions.Apply(new[] { region.ToRedactionAreaForTest() });

        Assert.DoesNotContain("Secret marker", redacted.Read.Text(), StringComparison.Ordinal);
        Assert.Empty(redacted.Read.AnnotationsBySubtype("Highlight"));
    }

    private static PdfPageRegion RegionAround(IReadOnlyList<PdfTextSpan> spans) {
        double left = spans.Min(static span => span.X);
        double right = spans.Max(static span => span.X + Math.Max(1D, Math.Abs(span.Advance)));
        double bottom = spans.Min(static span => span.Y - span.FontSize * 0.3D);
        double top = spans.Max(static span => span.Y + span.FontSize * 0.9D);
        return new PdfPageRegion(1, left - 0.5D, bottom, right - left + 1D, top - bottom);
    }

    private static byte[] BuildRawTextPdf(string content, string additionalResources = "", string additionalObjects = "", string pageEntries = "", string firstBaseFont = "Helvetica") {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] " + pageEntries + " /Resources << /Font << /F1 5 0 R /F2 6 0 R >> " + additionalResources + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /" + firstBaseFont + " >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj\n");
        WriteAscii(output, additionalObjects);
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 8 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

}

internal static class PdfTextEditorTestExtensions {
    internal static PdfRedactionArea ToRedactionAreaForTest(this PdfPageRegion region) =>
        new PdfRedactionArea(region.PageNumber, region.X, region.Y, region.Width, region.Height);
}
