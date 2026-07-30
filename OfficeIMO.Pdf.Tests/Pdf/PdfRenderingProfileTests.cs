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
    public void OverlayPreservesCallerNamedFamilyWhenProfileNameCollides() {
        byte[] callerData = ManagedTextShapingTestAssets.CreateFont('A');
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Shared", callerData));
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Shared", ManagedTextShapingTestAssets.CreateFont('B'));

        options.UseRenderingProfile(
            new OfficeRenderingProfile("named-collision", profileFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Equal(callerData, options.NamedFontFamilies["Shared"].Regular);
        Assert.False(options.TryGetRenderingProfileFamilyFallbacks(
            "Shared",
            out _));
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void ReplacingProfileOwnedNamedFamilyRemovesOldActiveFallbackBytes() {
        byte[] profileData = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] callerData = ManagedTextShapingTestAssets.CreateFont('B');
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Shared", profileData)
            .AddFallbackFamily("Shared");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("profile", profileFonts));

        options.RegisterNamedFontFamily(
            new PdfEmbeddedFontFamily("Shared", callerData));

        Assert.Equal(callerData, options.NamedFontFamilies["Shared"].Regular);
        Assert.Null(options.EmbeddedFontFallbacks);
        Assert.Null(options.GetEffectiveRenderingProfileDeclaredFallbacks(
            bold: false,
            italic: false));
    }

    [Fact]
    public void ReplacingProfileFallbackSetPreservesTheNewCallerCandidate() {
        byte[] callerData = ManagedTextShapingTestAssets.CreateFont('B');
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Shared", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddFallbackFamily("Shared");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile(
                "profile-fallback",
                profileFonts));

        options.RegisterEmbeddedFontFallbacks(
            new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Shared", callerData)
            }));

        PdfEmbeddedFontFallbackCandidate candidate = Assert.Single(
            options.EmbeddedFontFallbacks!.Candidates);
        Assert.Equal(callerData, candidate.DataSnapshot);
        Assert.Equal(callerData, options.NamedFontFamilies["Shared"].Regular);
        Assert.True(options.EmbeddedFontFallbacks.PlanText("B").IsFullyCovered);
    }

    [Fact]
    public void OverlayRefreshesFallbackBytesOwnedByEarlierProfile() {
        byte[] firstData = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] secondData = ManagedTextShapingTestAssets.CreateFont('B');
        var firstFonts = new OfficeFontFaceCollection()
            .Add("Shared", firstData)
            .AddFallbackFamily("Shared");
        var secondFonts = new OfficeFontFaceCollection()
            .Add("Shared", secondData)
            .AddFallbackFamily("Shared");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("first", firstFonts));

        options.UseRenderingProfile(
            new OfficeRenderingProfile("second", secondFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Equal(secondData, options.NamedFontFamilies["Shared"].Regular);
        Assert.Equal(
            secondData,
            Assert.Single(options.EmbeddedFontFallbacks!.Candidates).DataSnapshot);
    }

    [Fact]
    public void OverlayRefreshesInheritedFallbackWithoutRedeclaration() {
        byte[] firstData = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] secondData = ManagedTextShapingTestAssets.CreateFont('B');
        var firstFonts = new OfficeFontFaceCollection()
            .Add("Shared", firstData)
            .AddFallbackFamily("Shared");
        var secondFonts = new OfficeFontFaceCollection()
            .Add("Shared", secondData);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("first", firstFonts));

        options.UseRenderingProfile(
            new OfficeRenderingProfile("second", secondFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Equal(
            secondData,
            Assert.Single(options.EmbeddedFontFallbacks!.Candidates).DataSnapshot);
    }

    [Fact]
    public void OverlayAppendsNewFallbackFamiliesAfterInheritedOrder() {
        var firstFonts = new OfficeFontFaceCollection()
            .Add("First", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddFallbackFamily("First");
        var secondFonts = new OfficeFontFaceCollection()
            .Add("Second", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddFallbackFamily("Second");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("first", firstFonts));

        options.UseRenderingProfile(
            new OfficeRenderingProfile("second", secondFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Equal(
            new[] { "First", "Second" },
            options.EmbeddedFontFallbacks!.FontFamilyNames);
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
    public void OverlayPrependsNewOverlappingRangeScopedPlannerCandidate() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var aThroughB = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'B')
        });
        byte[] firstData = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] newestData = ManagedTextShapingTestAssets.CreateFont('A', 'B');
        var first = new OfficeFontFaceCollection()
            .Add("Scoped", firstData, OfficeFontStyle.Regular, onlyA);
        var second = new OfficeFontFaceCollection()
            .Add("Scoped", newestData, OfficeFontStyle.Regular, aThroughB);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("first", first))
            .UseRenderingProfile(
                new OfficeRenderingProfile("second", second),
                OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.TryGetRenderingProfileFamilyFallbacks(
            "Scoped",
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfEmbeddedFontFallbackSet planner =
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks);
        PdfTextFallbackSegment segment = Assert.Single(planner.PlanText("A").Segments);

        Assert.Equal(newestData, planner.Candidates[segment.FontIndex].DataSnapshot);
    }

    [Fact]
    public void RequestedFamilyListCombinesEveryScopedFamily() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var onlyB = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('B', 'B')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add("First", ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular, onlyA)
            .Add("Second", ManagedTextShapingTestAssets.CreateFont('B'), OfficeFontStyle.Regular, onlyB);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("multiple-scoped", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "First, Second",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackPlan plan = Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("AB");

        Assert.True(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);
    }

    [Fact]
    public void RequestedFamilyListIncludesUnrestrictedAndScopedFamiliesInOrder() {
        var onlyB = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('B', 'B')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add("Primary", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('B'),
                OfficeFontStyle.Regular,
                onlyB);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("mixed-families", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Primary, Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackPlan plan = Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("AB");

        Assert.True(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);
        Assert.Equal("Primary", plan.Segments[0].FontName);
        Assert.Equal(fonts.Faces[1].ResourceFamilyName, plan.Segments[1].FontName);
    }

    [Fact]
    public void RequestedRangeFamilyListPreservesFamilyPriorityBeforeStyle() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "First",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add(
                "Second",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Bold,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("styled-family-list", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "First, Second",
            bold: true,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfTextFallbackSegment segment = Assert.Single(
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
                .PlanText("A")
                .Segments);

        Assert.Equal(fonts.Faces[0].ResourceFamilyName, segment.FontName);
    }

    [Fact]
    public void UnrestrictedOnlyOverlayRefreshesExistingScopedFamilyPlanner() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var first = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular, onlyA)
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('B'));
        var second = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('C'));
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
        PdfEmbeddedFontFallbackSet planner =
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks);

        Assert.True(planner.PlanText("AC").IsFullyCovered);
        Assert.False(planner.PlanText("B").IsFullyCovered);
        Assert.Equal(
            second.Faces[0].Data,
            Assert.Single(planner.Candidates, candidate => candidate.FontName == "Scoped")
                .DataSnapshot);
    }

    [Fact]
    public void RegularFallbackUsesAvailableBoldFaceWhenRegularIsAbsent() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add("Primary", ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular, onlyA)
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('B'), OfficeFontStyle.Bold)
            .AddFallbackFamily("Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("bold-only-fallback", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Primary",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));

        Assert.True(Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("B").IsFullyCovered);
    }

    [Fact]
    public void EffectiveProfilePlannerRegistersSlotFallbackAsNamedFamily() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular, onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("slot-fallback", fonts))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate(
                        "Slot Fallback",
                        ManagedTextShapingTestAssets.CreateFont('B'))
                },
                new[] { PdfStandardFont.Helvetica }));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfEmbeddedFontFallbackSet planner =
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks);

        Assert.True(options.HasNamedFontFamily("Slot Fallback"));
        Assert.True(planner.PlanText("AB").IsFullyCovered);
    }

    [Fact]
    public void LateSlotFallbackCollisionUsesDistinctNamedFallbackBytes() {
        byte[] slotData = ManagedTextShapingTestAssets.CreateFont('A');
        var fonts = new OfficeFontFaceCollection()
            .Add("Shared", ManagedTextShapingTestAssets.CreateFont('B'));
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile(
                "late-slot-collision",
                fonts))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate("Shared", slotData)
                },
                new[] { PdfStandardFont.Helvetica }));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Shared",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfEmbeddedFontFallbackSet planner =
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks);
        PdfTextFallbackSegment segment = Assert.Single(
            planner.PlanText("A").Segments);
        PdfEmbeddedFontFallbackCandidate selected =
            planner.Candidates[segment.FontIndex];

        Assert.NotEqual("Shared", selected.FontName);
        Assert.Equal(slotData, selected.DataSnapshot);
        Assert.Equal(
            slotData,
            options.NamedFontFamilies[selected.FontName].Regular);
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
    public void EncodingPreflightResetsRangeScopedWhitespaceContextAtTextBoundaries() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A', 0x2003),
                OfficeFontStyle.Regular,
                onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("scoped-boundary", fonts));

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics =
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] { TextRun.Normal("A\n\u2003", fontFamily: "Scoped") },
                options,
                PdfStandardFont.Helvetica,
                "profile boundary preflight");

        Assert.Single(diagnostics);
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
                OfficeFontStyle.Bold,
                onlyA)
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
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
    public void ReplaceProfileDoesNotExcludeUnrestrictedFaceReusingCallerFamilyName() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                "Shared",
                ManagedTextShapingTestAssets.CreateFont('A')));
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Shared",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add("Shared", ManagedTextShapingTestAssets.CreateFont('B'));

        options.UseRenderingProfile(new OfficeRenderingProfile("replacement", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Shared",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        Assert.True(Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("B").IsFullyCovered);
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
    public void RangeScopedPlannerPrefersNewestOverlappingFace() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var first = ManagedTextShapingTestAssets.CreateFont('A');
        var newest = ManagedTextShapingTestAssets.CreateFont('A', 'B');
        var fonts = new OfficeFontFaceCollection()
            .Add("Scoped", first, OfficeFontStyle.Regular, onlyA)
            .Add("Scoped", newest, OfficeFontStyle.Regular, onlyA);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("newest-scoped", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "Scoped",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));
        PdfEmbeddedFontFallbackSet planner =
            Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks);
        PdfTextFallbackSegment segment = Assert.Single(planner.PlanText("A").Segments);

        Assert.Equal(newest, planner.Candidates[segment.FontIndex].DataSnapshot);
    }

    [Fact]
    public void RangeScopedPlannerResolvesOfficeFamilyListsInOrder() {
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
            .UseRenderingProfile(new OfficeRenderingProfile("family-list", fonts));

        Assert.True(options.TryGetEffectiveRenderingProfileFallbacks(
            "\"Missing\", \"Scoped\", \"Backup\"",
            bold: false,
            italic: false,
            out PdfEmbeddedFontFallbackSet? fallbacks));

        Assert.True(Assert.IsType<PdfEmbeddedFontFallbackSet>(fallbacks)
            .PlanText("A").IsFullyCovered);
    }

    [Fact]
    public void CallerNamedFamilyWinsBeforeLaterProfileFamilyInOfficeList() {
        var options = new PdfOptions {
            CompressContentStreams = false
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            "Caller",
            ManagedTextShapingTestAssets.CreateFont(' ', 'A')));
        var fonts = new OfficeFontFaceCollection()
            .Add("Profile", ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B'));
        options.UseRenderingProfile(
            new OfficeRenderingProfile("caller-priority", fonts),
            OfficeRenderingProfileApplyMode.Overlay);

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph
                .FontFamily("Caller, Profile")
                .Text("A"))
            .ToBytes();
        string raw = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Contains("/BaseFont /Caller-Regular", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Profile-Regular", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void EncodingPreflightCombinesCallerAndProfileFamiliesInOfficeList() {
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                "Caller",
                ManagedTextShapingTestAssets.CreateFont('A')));
        var fonts = new OfficeFontFaceCollection()
            .Add("Profile", ManagedTextShapingTestAssets.CreateFont('B'));
        options.UseRenderingProfile(
            new OfficeRenderingProfile("caller-profile-preflight", fonts),
            OfficeRenderingProfileApplyMode.Overlay);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics =
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] {
                    TextRun.Normal("AB", fontFamily: "Caller, Profile")
                },
                options,
                PdfStandardFont.Helvetica,
                "mixed family preflight");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void CallerAndProfileFallbackDoNotSplitCombiningGraphemeAcrossFonts() {
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                "Caller",
                ManagedTextShapingTestAssets.CreateFont(' ', 'a')));
        var markRange = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('\u0301', '\u0301')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Marks",
                ManagedTextShapingTestAssets.CreateFont('\u0301'),
                OfficeFontStyle.Regular,
                markRange)
            .AddFallbackFamily("Marks");
        options.UseRenderingProfile(
            new OfficeRenderingProfile("mark-fallback", fonts),
            OfficeRenderingProfileApplyMode.Overlay);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics =
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] {
                    TextRun.Normal("a\u0301", fontFamily: "Caller, Marks")
                },
                options,
                PdfStandardFont.Helvetica,
                "mixed grapheme preflight");

        Assert.Single(diagnostics);

        PdfTextEncodingPreflightException exception = Assert.Throws<PdfTextEncodingPreflightException>(() =>
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph
                    .FontFamily("Caller")
                    .Text("a\u0301"))
                .ToBytes());

        PdfTextEncodingDiagnostic writerDiagnostic = Assert.Single(exception.TextEncodingDiagnostics);
        Assert.Equal(0, writerDiagnostic.Index);
    }

    [Fact]
    public void WinAnsiAndRegisteredFallbackDoNotSplitCombiningGraphemeAcrossFonts() {
        var options = new PdfOptions()
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate(
                    "Marks",
                    ManagedTextShapingTestAssets.CreateFont('\u0301'))
            }));

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics =
            PdfTextDiagnostics.AnalyzeGeneratedText(
                "a\u0301",
                options,
                PdfStandardFont.Helvetica,
                "ordinary fallback grapheme preflight");

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(0, diagnostic.Index);
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
    public void BoldItalicFallbackUsesRegularBeforePartialStyles() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add(
                "Fallback",
                ManagedTextShapingTestAssets.CreateFont('B'),
                OfficeFontStyle.Bold)
            .AddFallbackFamily("Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("bold-italic", fonts));

        PdfEmbeddedFontFallbackSet planner = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.GetEffectiveRenderingProfileDeclaredFallbacks(
                bold: true,
                italic: true));

        Assert.True(planner.PlanText("A").IsFullyCovered);
        Assert.False(planner.PlanText("B").IsFullyCovered);
    }

    [Fact]
    public void FallbackWithoutExactOrRegularFaceUsesReverseRegistrationOrder() {
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "Fallback",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Bold)
            .Add(
                "Fallback",
                ManagedTextShapingTestAssets.CreateFont('B'),
                OfficeFontStyle.Italic)
            .AddFallbackFamily("Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("registration-order", fonts));

        PdfEmbeddedFontFallbackSet planner = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.GetEffectiveRenderingProfileDeclaredFallbacks(
                bold: false,
                italic: false));

        Assert.True(planner.PlanText("B").IsFullyCovered);
        Assert.False(planner.PlanText("A").IsFullyCovered);
    }

    [Fact]
    public void DiagnosticsUseStyledGlobalFallbackWhenFamilyPlannerIsUnavailable() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('\u0416'))
            .Add(
                "Fallback",
                ManagedTextShapingTestAssets.CreateFont('\u0419'),
                OfficeFontStyle.Bold)
            .AddFallbackFamily("Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("diagnostic-style", fonts));

        IReadOnlyList<PdfTextEncodingDiagnostic> covered =
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] {
                    TextRun.Bolded("\u0419", fontFamily: "Missing")
                },
                options,
                PdfStandardFont.Helvetica,
                "styled global fallback");
        IReadOnlyList<PdfTextEncodingDiagnostic> uncovered =
            PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
                new[] {
                    TextRun.Bolded("\u0416", fontFamily: "Missing")
                },
                options,
                PdfStandardFont.Helvetica,
                "styled global fallback");

        Assert.Empty(covered);
        Assert.NotEmpty(uncovered);
    }

    [Fact]
    public void ResourceScopedFallbackStaysScopedAcrossOverlay() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var initialFonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('A'),
                OfficeFontStyle.Regular,
                onlyA);
        initialFonts.AddFallbackFamily(initialFonts.Faces[0].ResourceFamilyName);
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("resource-scoped", initialFonts));
        var onlyC = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('C', 'C')
        });
        var overlayFonts = new OfficeFontFaceCollection()
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFont('C'),
                OfficeFontStyle.Regular,
                onlyC);

        options.UseRenderingProfile(
            new OfficeRenderingProfile("resource-overlay", overlayFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        PdfEmbeddedFontFallbackSet planner = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.GetEffectiveRenderingProfileDeclaredFallbacks(
                bold: false,
                italic: false));
        Assert.True(planner.PlanText("A").IsFullyCovered);
        Assert.False(planner.PlanText("C").IsFullyCovered);
    }

    [Fact]
    public void OverlaySlotFallbackCollisionStillRegistersProfileNamedFamily() {
        byte[] slotData = ManagedTextShapingTestAssets.CreateFont('A');
        var options = new PdfOptions()
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate(
                        "Shared",
                        slotData)
                },
                new[] { PdfStandardFont.Helvetica }));
        var fonts = new OfficeFontFaceCollection()
            .Add("Shared", ManagedTextShapingTestAssets.CreateFont('B'));

        options.UseRenderingProfile(
            new OfficeRenderingProfile("slot-collision", fonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.HasNamedFontFamily("Shared"));
        Assert.True(options.NamedFontFamilies["Shared"].Regular
            .SequenceEqual(fonts.Faces[0].Data));
        PdfEmbeddedFontFallbackCandidate preserved = Assert.Single(
            options.EmbeddedFontFallbacks!.Candidates,
            candidate => candidate.FontName != "Shared");
        Assert.Equal(slotData, preserved.DataSnapshot);
        Assert.Equal(slotData, options.NamedFontFamilies[preserved.FontName].Regular);
        Assert.True(options.EmbeddedFontFallbacks.PlanText("A").IsFullyCovered);
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
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void ClearingNamedFamiliesPreservesCallerOwnedSlotFallbackPlanner() {
        var slotFallbacks = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate(
                    "Caller slot",
                    ManagedTextShapingTestAssets.CreateFont('A'))
            },
            new[] { PdfStandardFont.Helvetica });
        var options = new PdfOptions()
            .RegisterEmbeddedFontFallbacks(slotFallbacks);

        options.ClearNamedFontFamilies();

        Assert.Equal(
            new[] { PdfStandardFont.Helvetica },
            options.EmbeddedFontFallbacks?.FontSlots);
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
    public void FallbackPlannerKeepsCombiningGraphemeWithinOneFont() {
        var baseRange = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('a', 'a')
        });
        var markRange = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange(0x0301, 0x0301)
        });
        PdfTextFallbackPlan splitCandidates = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate(
                    "Base",
                    ManagedTextShapingTestAssets.CreateFont('a'),
                    baseRange),
                new PdfEmbeddedFontFallbackCandidate(
                    "Mark",
                    ManagedTextShapingTestAssets.CreateFont(0x0301),
                    markRange)
            })
            .PlanText("a\u0301");
        PdfTextFallbackPlan completeCandidate = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate(
                    "Complete",
                    ManagedTextShapingTestAssets.CreateFont('a', 0x0301))
            })
            .PlanText("a\u0301");

        Assert.False(splitCandidates.IsFullyCovered);
        Assert.Empty(splitCandidates.Segments);
        Assert.Single(splitCandidates.Diagnostics);
        Assert.True(completeCandidate.IsFullyCovered);
        Assert.Equal("a\u0301", Assert.Single(completeCandidate.Segments).Text);
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
    public void RenderingProfileFamilyCapacityIsValidatedBeforeOptionsAreMutated() {
        var originalProvider = new DecliningTextShapingProvider();
        var replacementProvider = OfficeManagedTextShapingProvider.Instance;
        var options = new PdfOptions {
            TextShapingProvider = originalProvider,
            Language = "en"
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            "Existing",
            ManagedTextShapingTestAssets.CreateFont('A')));
        var fonts = new OfficeFontFaceCollection();
        for (int index = 0; index <= PdfOptions.MaximumNamedFontFamilies; index++) {
            fonts.Add(
                $"Family {index}",
                ManagedTextShapingTestAssets.CreateFont('A'));
        }

        Assert.Throws<InvalidOperationException>(() =>
            options.UseRenderingProfile(new OfficeRenderingProfile(
                "too-many-families",
                fonts,
                replacementProvider,
                "pl")));

        Assert.Same(originalProvider, options.TextShapingProvider);
        Assert.Equal("en", options.Language);
        Assert.True(options.HasNamedFontFamily("Existing"));
        Assert.Single(options.NamedFontFamilies);
    }

    [Fact]
    public void RenderingProfileCapacityIncludesPromotedCompatibilityFallbacks() {
        var originalProvider = new DecliningTextShapingProvider();
        var options = new PdfOptions {
            TextShapingProvider = originalProvider,
            Language = "en"
        };
        byte[] data = ManagedTextShapingTestAssets.CreateFont('A');
        for (int index = 0; index < PdfOptions.MaximumNamedFontFamilies - 1; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                $"Existing {index}",
                data));
        }
        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Slot", data) },
            new[] { PdfStandardFont.Helvetica }));
        var fonts = new OfficeFontFaceCollection().Add("Profile", data);

        Assert.Throws<InvalidOperationException>(() =>
            options.UseRenderingProfile(
                new OfficeRenderingProfile(
                    "slot-capacity",
                    fonts,
                    OfficeManagedTextShapingProvider.Instance,
                    "pl"),
                OfficeRenderingProfileApplyMode.Overlay));

        Assert.Same(originalProvider, options.TextShapingProvider);
        Assert.Equal("en", options.Language);
        Assert.Equal(
            PdfOptions.MaximumNamedFontFamilies - 1,
            options.NamedFontFamilies.Count);
        Assert.False(options.EmbeddedFontFallbacks!.UsesNamedFontFamilies);
    }

    [Fact]
    public void CallerFallbackRetainsPromotedProfileFamilyWithIdenticalBytes() {
        byte[] data = ManagedTextShapingTestAssets.CreateFont('A');
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Shared", data)
            .AddFallbackFamily("Shared");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile(
                "shared-profile",
                profileFonts));

        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Shared", data) },
            new[] { PdfStandardFont.Helvetica }));

        PdfEmbeddedFontFallbackSet fallback = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.EmbeddedFontFallbacks);
        Assert.True(fallback.UsesNamedFontFamilies);
        Assert.True(fallback.PlanText("A").IsFullyCovered);
    }

    [Fact]
    public void FontConfigurationStateIncludesFallbackUnicodeRanges() {
        byte[] data = ManagedTextShapingTestAssets.CreateFont('A', 'B');
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var onlyB = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('B', 'B')
        });
        var options = new PdfOptions {
            EmbeddedFontFallbacks = new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate("Scoped", data, onlyA)
                },
                new[] { PdfStandardFont.Helvetica })
        };
        long first = options.FontConfigurationState;

        options.EmbeddedFontFallbacks = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Scoped", data, onlyB)
            },
            new[] { PdfStandardFont.Helvetica });

        Assert.NotEqual(first, options.FontConfigurationState);
    }

    [Fact]
    public void AssigningEmbeddedFallbacksReleasesProfileDeclaredFallbacks() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Profile Fallback", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddFallbackFamily("Profile Fallback");
        var options = new PdfOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("profile", fonts));

        options.EmbeddedFontFallbacks = null;

        Assert.Null(options.EmbeddedFontFallbacks);
        Assert.Null(options.GetEffectiveRenderingProfileDeclaredFallbacks(
            bold: false,
            italic: false));
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
        Assert.False(word.HasExplicitPdfFontConfiguration);
        Assert.False(word.CloneForConversion().HasExplicitPdfFontConfiguration);

        var explicitlyConfiguredWord = new OfficeIMO.Word.Pdf.WordPdfSaveOptions {
            PdfOptions = new PdfOptions()
        }.UseRenderingProfile(profile);
        Assert.True(explicitlyConfiguredWord.HasExplicitPdfFontConfiguration);
        Assert.True(explicitlyConfiguredWord
            .CloneForConversion()
            .HasExplicitPdfFontConfiguration);

        word.PdfOptions!.DefaultFontSize = 17D;
        Assert.True(word.HasExplicitPdfFontConfiguration);
        excel.PdfOptions!.PageSize = new PageSize(300, 400);
    }

    [Fact]
    public void OfficePdfAdaptersRejectInvalidProfilesWithoutCreatingPdfOptions() {
        var excel = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions();
        var word = new OfficeIMO.Word.Pdf.WordPdfSaveOptions();
        var profile = new OfficeRenderingProfile("invalid-mode");
        var invalidMode = (OfficeRenderingProfileApplyMode)int.MaxValue;

        Assert.Throws<ArgumentNullException>(() => excel.UseRenderingProfile(null!));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            excel.UseRenderingProfile(profile, invalidMode));
        Assert.Null(excel.PdfOptions);

        Assert.Throws<ArgumentNullException>(() => word.UseRenderingProfile(null!));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            word.UseRenderingProfile(profile, invalidMode));
        Assert.Null(word.PdfOptions);
    }

    [Fact]
    public void ExcelShapingOnlyProfileDoesNotBecomeExplicitFontConfiguration() {
        var options = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        System.Reflection.PropertyInfo state = typeof(OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions)
            .GetProperty(
                "HasExplicitPdfFontConfiguration",
                System.Reflection.BindingFlags.Instance
                | System.Reflection.BindingFlags.NonPublic)!;

        Assert.False(Assert.IsType<bool>(state.GetValue(options)));

        options.PdfOptions!.DefaultFontSize = 13;

        Assert.True(Assert.IsType<bool>(state.GetValue(options)));
    }

    [Fact]
    public void WordShapingOnlyProfileTreatsEqualFontSizeAssignmentAsExplicit() {
        var options = new OfficeIMO.Word.Pdf.WordPdfSaveOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));

        Assert.False(options.HasExplicitPdfFontConfiguration);

        options.PdfOptions!.DefaultFontSize = 11D;

        Assert.True(options.HasExplicitPdfFontConfiguration);
        Assert.True(options.CloneForConversion().HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void ExcelRenderingProfilePreservesWorksheetPageSizeUntilCallerOverridesIt() {
        using var workbook = OfficeIMO.Excel.ExcelDocument.Create(new MemoryStream());
        OfficeIMO.Excel.ExcelSheet sheet = workbook.AddWorksheet("Profile");
        sheet.CellValue(1, 1, "Profile page size");
        sheet.SetPaperSize(OfficeIMO.Excel.ExcelPaperSize.A3);
        var options = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions()
            .UseRenderingProfile(OfficeRenderingProfile.Managed);

        byte[] worksheetSized = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdf(
            workbook,
            options);
        using (var pdf = UglyToad.PdfPig.PdfDocument.Open(worksheetSized)) {
            UglyToad.PdfPig.Content.Page page = pdf.GetPage(1);
            Assert.InRange((double)page.Width, 841D, 843D);
            Assert.InRange((double)page.Height, 1190D, 1192D);
        }

        options.PdfOptions!.PageSize = new PageSize(300, 400);
        byte[] explicitlySized = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdf(
            workbook,
            options);
        using var explicitPdf = UglyToad.PdfPig.PdfDocument.Open(explicitlySized);
        UglyToad.PdfPig.Content.Page explicitPage = explicitPdf.GetPage(1);
        Assert.InRange((double)explicitPage.Width, 299.9D, 300.1D);
        Assert.InRange((double)explicitPage.Height, 399.9D, 400.1D);
    }

    [Fact]
    public void ExcelRenderingProfileHonorsEqualValueLetterPageSizeAssignment() {
        using var workbook = OfficeIMO.Excel.ExcelDocument.Create(new MemoryStream());
        OfficeIMO.Excel.ExcelSheet sheet = workbook.AddWorksheet("Profile");
        sheet.CellValue(1, 1, "Explicit Letter page size");
        sheet.SetPaperSize(OfficeIMO.Excel.ExcelPaperSize.A3);
        var options = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions()
            .UseRenderingProfile(OfficeRenderingProfile.Managed);

        options.PdfOptions!.PageSize = new PageSize(612, 792);
        byte[] explicitlySized = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdf(
            workbook,
            options);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(explicitlySized);
        UglyToad.PdfPig.Content.Page page = pdf.GetPage(1);
        Assert.InRange((double)page.Width, 611.9D, 612.1D);
        Assert.InRange((double)page.Height, 791.9D, 792.1D);
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

    [Fact]
    public void PowerPointShapingOnlyProfileDoesNotBecomeExplicitFontConfiguration() {
        var options = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));

        Assert.False(options.HasExplicitPdfFontConfiguration);

        options.PdfOptions!.DefaultFont = PdfStandardFont.Courier;

        Assert.True(options.HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void PowerPointFontProfileRemainsExplicitFontConfiguration() {
        var profile = new OfficeRenderingProfile(
            "font-profile",
            new OfficeFontFaceCollection()
                .Add("Profile", ManagedTextShapingTestAssets.CreateFont('A')));
        var options = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
            .UseRenderingProfile(profile);

        Assert.True(options.HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void WordFontProfileRemainsExplicitFontConfiguration() {
        var profile = new OfficeRenderingProfile(
            "word-font-profile",
            new OfficeFontFaceCollection()
                .Add("Profile", ManagedTextShapingTestAssets.CreateFont('A')));
        var options = new OfficeIMO.Word.Pdf.WordPdfSaveOptions()
            .UseRenderingProfile(profile);

        Assert.True(options.HasExplicitPdfFontConfiguration);
        Assert.True(options.CloneForConversion().HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void WordFontProfileRemainsExplicitAfterFontlessOverlay() {
        var fontProfile = new OfficeRenderingProfile(
            "word-font-profile",
            new OfficeFontFaceCollection()
                .Add("Profile", ManagedTextShapingTestAssets.CreateFont('A')));
        var options = new OfficeIMO.Word.Pdf.WordPdfSaveOptions()
            .UseRenderingProfile(fontProfile);

        options.UseRenderingProfile(
            new OfficeRenderingProfile("shaping-only-overlay"),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.HasExplicitPdfFontConfiguration);
        Assert.True(options.CloneForConversion().HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void PowerPointFailedProfileDoesNotCreatePdfOptions() {
        var fonts = new OfficeFontFaceCollection();
        for (int index = 0; index <= PdfOptions.MaximumNamedFontFamilies; index++) {
            fonts.Add(
                $"Family {index}",
                ManagedTextShapingTestAssets.CreateFont('A'));
        }
        var options = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions();

        Assert.Throws<InvalidOperationException>(() =>
            options.UseRenderingProfile(new OfficeRenderingProfile(
                "too-many-families",
                fonts)));

        Assert.Null(options.PdfOptions);
        Assert.False(options.HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void PowerPointFontProfileCanBeReplacedByShapingOnlyProfile() {
        var fontProfile = new OfficeRenderingProfile(
            "font-profile",
            new OfficeFontFaceCollection()
                .Add("Profile", ManagedTextShapingTestAssets.CreateFont('A')));
        var options = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
            .UseRenderingProfile(fontProfile);

        options.UseRenderingProfile(
            new OfficeRenderingProfile("shaping-only"),
            OfficeRenderingProfileApplyMode.Replace);

        Assert.False(options.HasExplicitPdfFontConfiguration);
        Assert.False(options.CloneForConversion().HasExplicitPdfFontConfiguration);
    }

    private sealed class DecliningTextShapingProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => null;
    }
}
