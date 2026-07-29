using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfRenderingProfileTests {
    [Fact]
    public void SharedRenderingProfileConfiguresGeneratedPdfText() {
        var profile = new OfficeRenderingProfile(
            "managed-arabic",
            textShapingProvider: OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: " ar ");
        var options = new PdfOptions();

        PdfOptions returned = options.UseRenderingProfile(profile);

        Assert.Same(options, returned);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, options.TextShapingProvider);
        Assert.Equal("ar", options.Language);
    }

    [Fact]
    public void OverlayPreservesExistingProviderWhenProfileDeclinesToOwnIt() {
        var existing = new DecliningTextShapingProvider();
        var options = new PdfOptions {
            TextShapingProvider = existing,
            Language = "pl"
        };

        options.UseRenderingProfile(
            new OfficeRenderingProfile("fonts-only"),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Same(existing, options.TextShapingProvider);
        Assert.Equal("pl", options.Language);
    }

    [Fact]
    public void SharedRenderingProfileRegistersDeterministicFontsAndFallbacks() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Profile Sans", ManagedTextShapingTestAssets.CreateFont(' ', 'A'))
            .AddFallbackFamily("Profile Sans");
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("portable", fonts));

        Assert.True(options.HasNamedFontFamily("Profile Sans"));
        Assert.Equal(
            new[] { "Profile Sans" },
            options.EmbeddedFontFallbacks?.FontFamilyNames);
    }

    [Fact]
    public void SharedRenderingProfileDoesNotPromoteNamedFontsIntoUndeclaredFallbacks() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Named Only", ManagedTextShapingTestAssets.CreateFont('A'));
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("named-only", fonts));

        Assert.True(options.HasNamedFontFamily("Named Only"));
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void SharedRenderingProfilePreservesFallbackOrderAndUnicodeRanges() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "First",
                ManagedTextShapingTestAssets.CreateFont('A', 'B'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add("Second", ManagedTextShapingTestAssets.CreateFont('A', 'B'))
            .AddFallbackFamily("First")
            .AddFallbackFamily("Second");
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("ranged", fonts));

        PdfEmbeddedFontFallbackSet fallbacks = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.EmbeddedFontFallbacks);
        Assert.Equal(
            new[] {
                fonts.Faces[0].ResourceFamilyName,
                fonts.Faces[1].ResourceFamilyName
            },
            fallbacks.FontFamilyNames);
        PdfTextFallbackSegment segment = Assert.Single(fallbacks.PlanText("B").Segments);
        Assert.Equal(1, segment.FontIndex);
        Assert.Equal(fonts.Faces[1].ResourceFamilyName, segment.FontName);
    }

    [Fact]
    public void OverlayMergesDeclaredProfileFallbacksAfterExplicitPdfFallbacks() {
        var options = new PdfOptions()
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate(
                        "Existing Fallback",
                        ManagedTextShapingTestAssets.CreateFont('B'))
                }));
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Profile Sans", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddFallbackFamily("Profile Sans");

        options.UseRenderingProfile(
            new OfficeRenderingProfile("overlay", profileFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.HasNamedFontFamily("Profile Sans"));
        Assert.Equal(
            new[] { "Existing Fallback", "Profile Sans" },
            options.EmbeddedFontFallbacks?.FontFamilyNames);
    }

    [Fact]
    public void OverlayPreservesExistingFallbackBytesWhenProfileNameCollides() {
        byte[] existingData = ManagedTextShapingTestAssets.CreateFont('B');
        var options = new PdfOptions()
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate("Shared", existingData)
                }));
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Shared", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddFallbackFamily("Shared");

        options.UseRenderingProfile(
            new OfficeRenderingProfile("overlay-collision", profileFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Equal(existingData, options.NamedFontFamilies["Shared"].Regular);
        Assert.Equal(existingData, options.EmbeddedFontFallbacks?.Candidates[0].DataSnapshot);
    }

    [Fact]
    public void ProfileFallbackPlannerDoesNotReplaceStyledNamedFamily() {
        byte[] regular = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] bold = ManagedTextShapingTestAssets.CreateFont('B');
        var fonts = new OfficeFontFaceCollection()
            .Add("Styled", regular)
            .Add("Styled", bold, OfficeFontStyle.Bold)
            .AddFallbackFamily("Styled");
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("styled", fonts));

        PdfEmbeddedFontFamily family = options.NamedFontFamilies["Styled"];
        Assert.Equal(regular, family.Regular);
        Assert.Equal(bold, family.Bold);
    }

    [Fact]
    public void RangeScopedProfileFacesRemainAddressableByAuthoredFamily() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("scoped", fonts));

        Assert.True(options.TryGetRenderingProfileFamilyFallbacks(
            "Scoped",
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackSegment segment = Assert.Single(
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
                .PlanText("A")
                .Segments);
        Assert.Equal(fonts.Faces[0].ResourceFamilyName, segment.FontName);
        Assert.True(options.HasNamedFontFamily(segment.FontName));
    }

    [Fact]
    public void OverlayPreservesPriorRangeScopedFacesForAuthoredFamily() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var onlyB = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('B', 'B')
        });
        var first = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular, onlyA);
        var second = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('B'), OfficeFontStyle.Regular, onlyB);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("first", first))
            .UseRenderingProfile(
                new OfficeRenderingProfile("second", second),
                OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackPlan plan = Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("AB");

        Assert.True(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);
    }

    [Fact]
    public void RangeScopedFamilyPlannerIncludesUnrestrictedCatchAllFace() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('B'));
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("scoped-catch-all", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackPlan plan = Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("AB");

        Assert.True(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);
        Assert.Equal(fonts.Faces[0].ResourceFamilyName, plan.Segments[0].FontName);
        Assert.Equal("Scoped", plan.Segments[1].FontName);
    }

    [Fact]
    public void OverlayReplacesMatchingRangeScopedPlannerCandidate()
    {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var first = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular, onlyA);
        var second = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('B'), OfficeFontStyle.Regular, onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("first", first))
            .UseRenderingProfile(
                new OfficeRenderingProfile("second", second),
                OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.TryGetRenderingProfileFamilyFallbacks(
            "Scoped",
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfEmbeddedFontFallbackCandidate candidate = Assert.Single(
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks).Candidates);

        Assert.Equal(second.Faces[0].Data, candidate.DataSnapshot);
        Assert.False(Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("A").IsFullyCovered);
    }

    [Fact]
    public void RangeScopedPlannerKeepsWhitespaceWithAdjacentFont()
    {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont(' ', 'A'),
                OfficeFontStyle.Regular,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("scoped-whitespace", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackPlan plan = Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("A A");

        Assert.True(plan.IsFullyCovered);
        Assert.Single(plan.Segments);
        Assert.Equal("A A", plan.Segments[0].Text);
    }

    [Fact]
    public void RangeScopedProfileFacesSelectRequestedRunStyle() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Bold,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("styled-scoped", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: true,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackSegment segment = Assert.Single(
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
                .PlanText("A")
                .Segments);

        Assert.Equal(
            fonts.Faces.Single(face => face.Style == OfficeFontStyle.Bold).ResourceFamilyName,
            segment.FontName);
    }

    [Fact]
    public void RangeScopedPlannerPreservesStyledUnrestrictedFaces() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('B'),
                OfficeFontStyle.Bold);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("styled-catch-all", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: true,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackSegment segment = Assert.Single(
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
                .PlanText("B")
                .Segments);

        Assert.Equal("Scoped", segment.FontName);
        Assert.Equal(
            OfficeFontStyle.Bold,
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
                .Candidates[segment.FontIndex].Style);
    }

    [Fact]
    public void RequestedRangeFamilyCombinesWithDeclaredFallbacks() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add("Secondary", ManagedTextShapingTestAssets.CreateFont('B'))
            .AddFallbackFamily("Secondary");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("mixed-scoped", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackPlan plan = Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("AB");

        Assert.True(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);
        Assert.Equal(fonts.Faces[0].ResourceFamilyName, plan.Segments[0].FontName);
        Assert.Equal("Secondary", plan.Segments[1].FontName);
    }

    [Fact]
    public void DeclaredFallbackPlannerUsesRequestedRunStyle()
    {
        var onlyC = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('C', 'C')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Primary",
                ManagedTextShapingTestAssets.CreateFont('C'),
                OfficeFontStyle.Regular,
                onlyC)
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add(
                "Fallback",
                ManagedTextShapingTestAssets.CreateFont('B'),
                OfficeFontStyle.Bold)
            .AddFallbackFamily("Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("styled-fallback", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Primary",
            bold: true,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfEmbeddedFontFallbackSet planner =
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks);

        Assert.True(planner.PlanText("B").IsFullyCovered);
        Assert.False(planner.PlanText("A").IsFullyCovered);
    }

    [Fact]
    public void GlobalDeclaredFallbackPlannerUsesRequestedRunStyle() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add(
                "Fallback",
                ManagedTextShapingTestAssets.CreateFont('B'),
                OfficeFontStyle.Bold)
            .AddFallbackFamily("Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("global-styled-fallback", fonts));

        PdfEmbeddedFontFallbackSet planner = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.GetEffectiveRenderingProfileDeclaredFallbacks(
                bold: true,
                italic: false));

        Assert.True(planner.PlanText("B").IsFullyCovered);
        Assert.False(planner.PlanText("A").IsFullyCovered);
    }

    [Fact]
    public void DeclaredFallbackPlannerPreservesFamilyPriorityBeforeStyle() {
        var fonts = new OfficeFontFaceCollection()
            .Add("First", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add(
                "Second",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Bold)
            .AddFallbackFamily("First")
            .AddFallbackFamily("Second");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("family-priority", fonts));

        PdfEmbeddedFontFallbackSet planner = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.GetEffectiveRenderingProfileDeclaredFallbacks(
                bold: true,
                italic: false));
        PdfTextFallbackSegment segment = Assert.Single(planner.PlanText("A").Segments);

        Assert.Equal("First", segment.FontName);
    }

    [Fact]
    public void EncodingPreflightUsesRangeScopedAuthoredFamilyPlanner() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("preflight-scoped", fonts));

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics =
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] { TextRun.Normal("A", fontFamily: "Scoped") },
                options,
                PdfStandardFont.Helvetica,
                "profile preflight");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void ClearingNamedFamiliesAlsoClearsRangeScopedAuthoredFamilyPlanner() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("clear-scoped", fonts));

        options.ClearNamedFontFamilies();

        Assert.False(options.TryGetRenderingProfileFamilyFallbacks(
            "Scoped",
            out _));
    }

    [Fact]
    public void LatinLigatureFallbackRangesApplyToSourceCharacters() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('f', 'i', 0xFB01);
        var latin = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange(0x0000, 0x00FF)
        });
        var presentationOnly = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange(0xFB01, 0xFB01)
        });

        PdfTextFallbackPlan allowed = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Latin", font, latin) })
            .PlanText("fi", shapingMode: PdfTextShapingMode.LatinLigatures);
        PdfTextFallbackPlan rejected = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate(
                    "Presentation only",
                    font,
                    presentationOnly)
            })
            .PlanText("fi", shapingMode: PdfTextShapingMode.LatinLigatures);

        Assert.True(allowed.IsFullyCovered);
        Assert.False(rejected.IsFullyCovered);
    }

    [Fact]
    public void ReplaceClearsPreviouslyRegisteredPdfFontState() {
        var options = new PdfOptions()
            .UseFontFamily(new PdfEmbeddedFontFamily(
                "Existing Standard",
                ManagedTextShapingTestAssets.CreateFont('C')))
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                "Existing Family",
                ManagedTextShapingTestAssets.CreateFont('A')))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate(
                        "Existing Fallback",
                        ManagedTextShapingTestAssets.CreateFont('B'))
                }));

        options.UseRenderingProfile(new OfficeRenderingProfile("managed-only"));

        Assert.Empty(options.NamedFontFamilies);
        Assert.Empty(options.EmbeddedFonts);
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void SharedRenderingProfileFlowsThroughFirstPartyOfficePdfAdapters() {
        var profile = new OfficeRenderingProfile(
            "managed-polish",
            textShapingProvider: OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: "pl");

        var word = new OfficeIMO.Word.Pdf.WordPdfSaveOptions().UseRenderingProfile(profile);
        var excel = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions().UseRenderingProfile(profile);
        var powerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
            .UseRenderingProfile(profile);

        Assert.Same(OfficeManagedTextShapingProvider.Instance, word.PdfOptions?.TextShapingProvider);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, excel.PdfOptions?.TextShapingProvider);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, powerPoint.PdfOptions?.TextShapingProvider);
        Assert.Equal("pl", word.PdfOptions?.Language);
        Assert.Equal("pl", excel.PdfOptions?.Language);
        Assert.Equal("pl", powerPoint.PdfOptions?.Language);
    }

    [Fact]
    public void FontlessReplacementRestoresOfficeAdapterFontDiscovery() {
        static bool HasExplicitFontConfiguration(object options) =>
            (bool)(options.GetType().GetProperty(
                    "HasExplicitPdfFontConfiguration",
                    System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.NonPublic)
                ?.GetValue(options)
                ?? throw new InvalidOperationException(
                    "The PDF adapter font-configuration state was not found."));

        var configuredFonts = new OfficeFontFaceCollection()
            .Add("Configured", ManagedTextShapingTestAssets.CreateFont('A'));
        var configured = new OfficeRenderingProfile("configured", configuredFonts);
        var fontless = new OfficeRenderingProfile("fontless");

        var word = new OfficeIMO.Word.Pdf.WordPdfSaveOptions()
            .UseRenderingProfile(configured)
            .UseRenderingProfile(fontless);
        var excel = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions()
            .UseRenderingProfile(configured)
            .UseRenderingProfile(fontless);
        var powerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
            .UseRenderingProfile(configured)
            .UseRenderingProfile(fontless);

        Assert.False(HasExplicitFontConfiguration(word));
        Assert.False(HasExplicitFontConfiguration(excel));
        Assert.False(HasExplicitFontConfiguration(powerPoint));
    }

    [Fact]
    public void SharedRenderingProfileSurvivesOfficeAdapterCloningAndPdfGeneration() {
        OfficeRenderingProfile profile = OfficeRenderingProfile.Managed;
        byte[] wordPdf;
        using (var wordStream = new MemoryStream())
        using (OfficeIMO.Word.WordDocument word = OfficeIMO.Word.WordDocument.Create(wordStream)) {
            word.AddParagraph("Word profile proof");
            wordPdf = OfficeIMO.Word.Pdf.WordPdfConverterExtensions.ToPdf(
                word,
                new OfficeIMO.Word.Pdf.WordPdfSaveOptions().UseRenderingProfile(profile));
        }

        byte[] excelPdf;
        using (OfficeIMO.Excel.ExcelDocument excel =
            OfficeIMO.Excel.ExcelDocument.Create(new MemoryStream())) {
            excel.AddWorksheet("Profile").CellValue(1, 1, "Excel profile proof");
            excelPdf = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdf(
                excel,
                new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions().UseRenderingProfile(profile));
        }

        byte[] powerPointPdf;
        using (OfficeIMO.PowerPoint.PowerPointPresentation powerPoint =
            OfficeIMO.PowerPoint.PowerPointPresentation.Create(new MemoryStream())) {
            powerPoint.AddSlide().AddTextBoxPoints(
                "PowerPoint profile proof",
                24,
                24,
                240,
                40);
            powerPointPdf = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPdf(
                powerPoint,
                new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
                    .UseRenderingProfile(profile));
        }

        Assert.Contains("Word profile proof", PdfReadDocument.Open(wordPdf).ExtractText());
        Assert.Contains("Excel profile proof", PdfReadDocument.Open(excelPdf).ExtractText());
        Assert.Contains("PowerPoint profile proof", PdfReadDocument.Open(powerPointPdf).ExtractText());
    }

    private sealed class DecliningTextShapingProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => null;
    }
}
