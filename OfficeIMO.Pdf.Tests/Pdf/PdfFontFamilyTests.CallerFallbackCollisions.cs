using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData(false, false, 500)]
    [InlineData(true, false, 800)]
    [InlineData(false, true, 600)]
    [InlineData(true, true, 900)]
    public void CallerNamedFallbackPreservesAuthoredFamilyFaceAdvances(bool bold, bool italic, int advance) {
        var family = CreateCallerCollisionFamily();
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            // The same regular bytes are enough to expose replacement of the other faces.
            new PdfEmbeddedFontFallbackCandidate(" Caller Family ", family.Regular)
        });
        var baselineOptions = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(family);
        var fallbackOptions = baselineOptions.Clone().RegisterEmbeddedFontFallbacks(fallbacks);

        byte[] Render(PdfCreateOptions options) => PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.FontFamily("Caller Family")
                .Bold(bold).Italic(italic).Text("AB"))
            .ToBytes();

        double baseline = GetCallerCollisionAdvance(Render(baselineOptions), "AB");
        Assert.InRange(baseline, advance / 50D - 0.001D, advance / 50D + 0.001D);
        double actual = GetCallerCollisionAdvance(Render(fallbackOptions), "AB");
        Assert.InRange(actual, baseline - 0.001D, baseline + 0.001D);
        Assert.Equal(family.Bold, fallbackOptions.NamedFontFamilies["Caller Family"].Bold);
        Assert.Equal(family.Italic, fallbackOptions.NamedFontFamilies["Caller Family"].Italic);
        Assert.Equal(family.BoldItalic, fallbackOptions.NamedFontFamilies["Caller Family"].BoldItalic);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CallerNamedFallbackOriginalPlannerRetainsItsFontWhenAuthoredFamilyNameCollides(bool tryPlan) {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9))
        });
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(CreateCallerCollisionFamily())
            .RegisterEmbeddedFontFallbacks(fallbacks);
        IReadOnlyList<PdfTextRun> runs;
        if (tryPlan) Assert.True(fallbacks.TryPlanTextRuns("A\u03a9", out runs));
        else runs = fallbacks.PlanTextRuns("A\u03a9");

        // The original reusable set, not only PdfOptions' private clone, can be
        // used by a caller. Both glyphs must keep its selected candidate metrics:
        // the authored family also covers A, but has a different advance.
        foreach (PdfCreateOptions configured in new[] { options, options.Clone() }) {
            byte[] pdf = PdfDocument.Create(configured)
                .Paragraph(paragraph => paragraph.Runs(runs))
                .ToBytes();
            Assert.InRange(GetCallerCollisionAdvance(pdf, "A\u03a9"), 5.999D, 6.001D);
        }
    }

    [Fact]
    public void CallerNamedFallbackRendersMissingGlyphWithoutReplacingCoveredStyledGlyph() {
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(CreateCallerCollisionFamily())
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Caller Family",
                    ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9))
            }));

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.FontFamily("Caller Family").Bold(true).Text("A\u03a9"))
            .ToBytes();

        // Covered A uses the authored bold face (8pt); Omega needs the fallback
        // face (3pt). Matching only extracted text would miss the lost bold face.
        Assert.InRange(GetCallerCollisionAdvance(pdf, "A\u03a9"), 10.999D, 11.001D);
        var glyphs = ReadCallerCollisionGlyphs(pdf);
        Assert.Equal(new[] { "A", "\u03a9" }, glyphs.Select(glyph => glyph.Text));
        Assert.InRange(glyphs[0].PaintedAdvance, 7.999D, 8.001D);
        Assert.InRange(glyphs[1].PaintedAdvance, 2.999D, 3.001D);
    }

    [Fact]
    public void CallerNamedFallbackRegistrationReusesRetainedResourceAcrossReplacementAndClear() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9);
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(CreateCallerCollisionFamily());
        PdfEmbeddedFontFallbackSet? latest = null;
        for (int index = 0; index < 65; index++) {
            latest = new PdfEmbeddedFontFallbackSet(new[] { new PdfEmbeddedFontFallbackCandidate("Caller Family", font) });
            options.RegisterEmbeddedFontFallbacks(latest);
            options.EmbeddedFontFallbacks = null;
        }

        Assert.Equal(2, options.NamedFontFamilies.Count);
        Assert.Null(options.EmbeddedFontFallbacks);
        IReadOnlyList<PdfTextRun> runs = latest!.PlanTextRuns("A\u03a9");
        Assert.Empty(PdfTextDiagnostics.AnalyzeGeneratedTextRuns(runs, options, PdfStandardFont.Helvetica));
        var document = PdfDocument.Create(options.Clone())
            .Paragraph(paragraph => paragraph.Runs(runs));
        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] pdf = document.ToBytes();
        Assert.InRange(GetCallerCollisionAdvance(pdf, "A\u03a9"), 5.999D, 6.001D);
        Assert.Equal(CreateCallerCollisionFamily().Bold, options.NamedFontFamilies["Caller Family"].Bold);
    }

    [Fact]
    public void CallerNamedFallbackRetainsPlannedFontAfterCallerClaimsItsRegisteredName() {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9))
        });
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }.RegisterEmbeddedFontFallbacks(fallbacks);
        options.RegisterNamedFontFamily(CreateCallerCollisionFamily());

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(fallbacks.PlanTextRuns("A\u03a9")))
            .Paragraph(paragraph => paragraph.FontFamily("Caller Family").Bold().Text("A\u03a9"))
            .ToBytes();
        IReadOnlyList<PdfTextGlyph> glyphs = ReadCallerCollisionGlyphs(pdf);
        Assert.Equal(new[] { "A", "\u03a9", "A", "\u03a9" }, glyphs.Select(glyph => glyph.Text));
        Assert.InRange(glyphs[0].PaintedAdvance, 2.999D, 3.001D);
        Assert.InRange(glyphs[1].PaintedAdvance, 2.999D, 3.001D);
        Assert.InRange(glyphs[2].PaintedAdvance, 7.999D, 8.001D);
        Assert.InRange(glyphs[3].PaintedAdvance, 2.999D, 3.001D);
    }

    [Fact]
    public void CallerNamedFallbackCollisionAtFamilyCapacityLeavesCallerFacesIntact() {
        var family = CreateCallerCollisionFamily();
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }.RegisterNamedFontFamily(family);
        for (int index = 1; index < 64; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Unrelated Family " + index, family.Regular));
        }
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family", family.Regular)
        });

        Assert.Throws<InvalidOperationException>(() => options.RegisterEmbeddedFontFallbacks(fallbacks));
        Assert.Equal(64, options.NamedFontFamilies.Count);
        Assert.Null(options.EmbeddedFontFallbacks);
        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.FontFamily("Caller Family").Bold().Text("AB"))
            .ToBytes();
        Assert.InRange(GetCallerCollisionAdvance(pdf, "AB"), 15.999D, 16.001D);
    }

    [Fact]
    public void CallerNamedFallbackClearNamedFamiliesDiscardsPlannedResourceBindings() {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9))
        });
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(CreateCallerCollisionFamily())
            .RegisterEmbeddedFontFallbacks(fallbacks);
        options.ClearNamedFontFamilies().RegisterNamedFontFamily(CreateCallerCollisionFamily());

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(fallbacks.PlanTextRuns("A")))
            .ToBytes();
        Assert.Null(options.EmbeddedFontFallbacks);
        Assert.Single(options.NamedFontFamilies);
        Assert.InRange(GetCallerCollisionAdvance(pdf, "A"), 4.999D, 5.001D);

        IReadOnlyList<PdfTextRun> missingRuns = fallbacks.PlanTextRuns("A\u03a9");
        PdfTextEncodingDiagnostic direct = Assert.Single(
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(missingRuns, options, PdfStandardFont.Helvetica));
        Assert.Equal("U+03A9", direct.CodePoint);
        var missingDocument = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(missingRuns));
        Assert.Equal("U+03A9", Assert.Single(missingDocument.AnalyzeTextEncoding()).CodePoint);
        Assert.Throws<PdfTextEncodingPreflightException>(() => missingDocument.ToBytes());
        Assert.Single(options.NamedFontFamilies);
    }

    [Fact]
    public void CallerNamedFallbackExplicitFontRegistrationReplacesRequestedNamedFamily() {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9))
        });
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }.RegisterNamedFontFamily(CreateCallerCollisionFamily());
        fallbacks.RegisterFonts(options);

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.FontFamily("Caller Family").Bold().Text("A\u03a9"))
            .ToBytes();
        Assert.Single(options.NamedFontFamilies);
        Assert.InRange(GetCallerCollisionAdvance(pdf, "A\u03a9"), 5.999D, 6.001D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CallerNamedFallbackExplicitPlanMappingKeepsTemplateFormattingAndNewFont(bool namedMapping) {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A'))
        });
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(CreateCallerCollisionFamily())
            .RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Explicit Mapped Font",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 'A')))
            .RegisterEmbeddedFontFallbacks(fallbacks);
        PdfTextRun template = fallbacks.PlanTextRuns("A")[0]
            .WithCharacterSpacing(global::OfficeIMO.OfficeLength.Points(1)).WithHorizontalTextScaling(50);
        PdfTextFallbackPlan plan = fallbacks.PlanText("A");
        IReadOnlyList<PdfTextRun> runs = namedMapping
            ? plan.ToNamedTextRuns(new[] { "Caller Family" }, template)
            : plan.ToTextRuns(new[] { PdfStandardFont.TimesRoman }, template);
        PdfTextRun expected = PdfTextRun.Normal("A",
                font: namedMapping ? null : PdfStandardFont.TimesRoman,
                fontFamily: namedMapping ? "Caller Family" : null)
            .WithCharacterSpacing(global::OfficeIMO.OfficeLength.Points(1)).WithHorizontalTextScaling(50);

        byte[] Render(IReadOnlyList<PdfTextRun> textRuns) => PdfDocument.Create(options.Clone())
            .Paragraph(paragraph => paragraph.Runs(textRuns)).ToBytes();
        double baseline = GetCallerCollisionAdvance(Render(new[] { expected }), "A");
        Assert.InRange(GetCallerCollisionAdvance(Render(runs), "A"), baseline - 0.001D, baseline + 0.001D);
    }

    [Fact]
    public void CallerNamedFallbackReplacesClonedProfileFamilyAndRetainsOwnershipAcrossOverlay() {
        PdfEmbeddedFontFamily originalFamily = CreateCallerCollisionFamily();
        var fonts = new OfficeFontFaceCollection()
            .Add(originalFamily.FamilyName, originalFamily.Regular)
            .Add(originalFamily.FamilyName, originalFamily.Bold!, OfficeFontStyle.Bold);
        var original = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .UseRenderingProfile(new OfficeRenderingProfile("profile-only-family", fonts));
        byte[] callerData = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9);
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate(" caller family ", callerData)
        });
        PdfCreateOptions options = original.Clone();

        options.EmbeddedFontFallbacks = fallbacks;

        Assert.Single(options.NamedFontFamilies);
        Assert.Equal(callerData, options.NamedFontFamilies[originalFamily.FamilyName].Regular);
        Assert.Equal(originalFamily.Regular, original.NamedFontFamilies[originalFamily.FamilyName].Regular);
        Assert.Equal(originalFamily.Bold, original.NamedFontFamilies[originalFamily.FamilyName].Bold);
        Assert.Null(original.EmbeddedFontFallbacks);
        var overlayFonts = new OfficeFontFaceCollection()
            .Add(originalFamily.FamilyName,
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 'A', 0x03A9))
            .AddFallbackFamily(originalFamily.FamilyName);
        options.UseRenderingProfile(new OfficeRenderingProfile("later-overlay", overlayFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Single(options.NamedFontFamilies);
        Assert.Equal(callerData, options.NamedFontFamilies[originalFamily.FamilyName].Regular);
        Assert.Equal(callerData, Assert.Single(options.EmbeddedFontFallbacks!.Candidates).DataSnapshot);
        byte[] pdf = PdfDocument.Create(options).Compose(document => document.Content(content => content
            .Paragraph(paragraph => paragraph.FontFamily(originalFamily.FamilyName).Text("A\u03a9"))
            .Paragraph(paragraph => paragraph.FontFamily(originalFamily.FamilyName).Bold().Text("A\u03a9"))
            .Paragraph(paragraph => paragraph.Runs(fallbacks.PlanTextRuns("A\u03a9")))))
            .ToBytes();
        Assert.InRange(GetCallerCollisionAdvance(pdf, "A\u03a9A\u03a9A\u03a9"), 17.999D, 18.001D);
    }

    [Fact]
    public void CallerNamedFallbackReusesProfileOwnedNameAtFamilyCapacity() {
        PdfEmbeddedFontFamily callerFamily = CreateCallerCollisionFamily();
        const string profileFamily = "Profile Family";
        var profileFonts = new OfficeFontFaceCollection()
            .Add(profileFamily, callerFamily.Regular)
            .AddFallbackFamily(profileFamily);
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .UseRenderingProfile(new OfficeRenderingProfile("profile-at-capacity", profileFonts));
        for (int index = 0; index < 63; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Other Family " + index,
                callerFamily.Regular, callerFamily.Bold));
        }
        byte[] replacement = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 0x03A9);

        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate(profileFamily, replacement)
        }));

        Assert.Equal(64, options.NamedFontFamilies.Count);
        Assert.Equal(replacement, options.NamedFontFamilies[profileFamily].Regular);
        Assert.Equal(callerFamily.Bold, options.NamedFontFamilies["Other Family 0"].Bold);
        Assert.Equal(callerFamily.Regular, options.NamedFontFamilies["Other Family 62"].Regular);
        Assert.Equal(profileFamily, Assert.Single(options.EmbeddedFontFallbacks!.Candidates).FontName);
        byte[] pdf = PdfDocument.Create(options).Compose(document => document.Content(content => content
            .Paragraph(paragraph => paragraph.FontFamily(profileFamily).Text("A\u03a9"))))
            .ToBytes();
        Assert.InRange(GetCallerCollisionAdvance(pdf, "A\u03a9"), 5.999D, 6.001D);
    }

    [Theory]
    [InlineData("bound-plan", "Fallback, Literal", "\"Fallback, Literal\"", "Fallback")]
    [InlineData("bound-plan", @"Fallback\zFace", "\"Fallback\\\\zFace\"", "FallbackzFace")]
    [InlineData("bound-plan", "Fallback Plain", "Fallback Plain", "Other Family")]
    [InlineData("mapped-plan", "Fallback, Literal", "\"Fallback, Literal\"", "Fallback")]
    [InlineData("mapped-plan", @"Fallback\zFace", "\"Fallback\\\\zFace\"", "FallbackzFace")]
    [InlineData("mapped-plan", "Fallback Plain", "Fallback Plain", "Other Family")]
    [InlineData("profile", "Fallback, Literal", "\"Fallback, Literal\"", "Fallback")]
    [InlineData("profile", @"Fallback\zFace", "\"Fallback\\\\zFace\"", "FallbackzFace")]
    [InlineData("profile", "Fallback Plain", "Fallback Plain", "Other Family")]
    public void NamedFallbackPublicPathsPreserveLiteralFamilyMetrics(
        string route, string familyName, string familyReference, string competingFamily) {
        byte[] familyData = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 900, ' ', 'A', 'B');
        byte[] competingData = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A', 'B');
        var options = new PdfCreateOptions { DefaultFontSize = OfficeLength.Points(10) };
        IReadOnlyList<PdfTextRun>? runs = null;
        if (route == "profile") {
            options.UseRenderingProfile(new OfficeRenderingProfile("literal-family", new OfficeFontFaceCollection()
                .Add(competingFamily, competingData).Add(familyName, familyData)));
        } else {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(competingFamily, competingData));
            var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate(familyName, familyData)
            });
            if (route == "bound-plan") {
                options.RegisterEmbeddedFontFallbacks(fallbacks);
                runs = fallbacks.PlanTextRuns("AB");
            } else {
                options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(familyName, familyData));
                runs = fallbacks.PlanText("AB").ToNamedTextRuns(new[] { familyName },
                    PdfTextRun.Normal("template", fontSize: OfficeLength.Points(10)));
            }
            // Plain names already represent one literal family and keep their public run value.
            if (familyName == familyReference) Assert.All(runs, run => Assert.Equal(familyName, run.FontFamily));
        }

        PdfDocument document = PdfDocument.Create(options.Clone());
        if (runs == null) document.Paragraph(paragraph => paragraph.FontFamily(familyReference).Text("AB"));
        else document.Paragraph(paragraph => paragraph.Runs(runs));
        byte[] pdf = document.ToBytes();
        Assert.InRange(GetCallerCollisionAdvance(pdf, "AB"), 17.999D, 18.001D);
        IReadOnlyList<PdfTextGlyph> glyphs = ReadCallerCollisionGlyphs(pdf);
        Assert.Equal(new[] { "A", "B" }, glyphs.Select(glyph => glyph.Text));
        Assert.All(glyphs, glyph => Assert.InRange(glyph.PaintedAdvance, 8.999D, 9.001D));
        Assert.Equal(familyData, options.NamedFontFamilies[familyName].Regular);
        Assert.Equal(competingData, options.NamedFontFamilies[competingFamily].Regular);
    }

    private static PdfEmbeddedFontFamily CreateCallerCollisionFamily() => new(
        "Caller Family",
        ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 'A', 'B'),
        ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 800, ' ', 'A', 'B'),
        ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, ' ', 'A', 'B'),
        ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 900, ' ', 'A', 'B'));

    private static double GetCallerCollisionAdvance(byte[] pdf, string expectedText) {
        var spans = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans()
            .Where(span => !string.IsNullOrEmpty(span.Text)).ToArray();
        Assert.Equal(expectedText, string.Concat(spans.Select(span => span.Text)));
        return spans.Sum(span => span.Advance);
    }

    private static IReadOnlyList<PdfTextGlyph> ReadCallerCollisionGlyphs(byte[] pdf) {
        var glyphs = new List<PdfTextGlyph>();
        foreach (PdfTextSpan span in PdfReadDocument.Open(pdf).Pages[0].GetTextSpans()) {
            if (string.IsNullOrEmpty(span.Text)) continue;
            Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> actual));
            glyphs.AddRange(actual);
        }
        return glyphs;
    }
}
