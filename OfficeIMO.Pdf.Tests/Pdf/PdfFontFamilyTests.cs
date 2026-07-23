using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontFamilyTests {
    [Fact]
    public void ExplicitDefaultFontResource_IsOmittedWhenOnlyNamedRunsUseText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var options = new PdfOptions {
            CompressContentStreams = false
        }
            .UseFontFamily(new PdfEmbeddedFontFamily("Configured Default", fontData))
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Visible Named", fontData));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph
                .FontFamily("Visible Named")
                .Text("Only the named family paints text."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("/BaseFont /ConfiguredDefault-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /VisibleNamed-Regular", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void NamedFontOnlyPage_DoesNotEmitUnusedStandardFontResources() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath == null) {
            return;
        }

        const string familyName = "Named Only";
        var options = new PdfOptions {
            CompressContentStreams = false
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(familyName, File.ReadAllBytes(fontPath)));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph
                .FontFamily(familyName)
                .Text("Only the registered family is used."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /NamedOnly-Regular", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Helvetica", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void NamedFontFamilies_RenderPageTextAndListMarkersAcrossFlowPaths() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string familyName = "Premium Named";
        byte[] fontData = File.ReadAllBytes(fontPath);
        var options = new PdfOptions {
            CompressContentStreams = false,
            ShowHeader = true,
            HeaderFormat = "Header",
            HeaderFontFamily = familyName,
            ShowPageNumbers = true,
            FooterFormat = "Footer",
            FooterFontFamily = familyName
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(familyName, fontData));
        var listStyle = new PdfListStyle {
            MarkerFontFamily = familyName
        };

        byte[] bytes = PdfDocument.Create(options)
            .RichNumbered(new[] { new PdfListItem("Top item", marker: "1.") }, style: listStyle)
            .Row(row => row.Column(100, column =>
                column.RichBullets(new[] { new PdfListItem("Column item", marker: "*") }, style: listStyle)))
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        foreach (string glyph in new[] { "H", "F", "1", "*" }) {
            Assert.Contains(
                page.Letters,
                letter =>
                    letter.Value == glyph &&
                    letter.FontName.Contains("PremiumNamed", StringComparison.OrdinalIgnoreCase));
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /PremiumNamed-Regular", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void NamedFontFamilies_RenderMoreThanThreeFamiliesOnOnePageWithoutSlotCollisions() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        string[] families = { "Named Alpha", "Named Bravo", "Named Charlie", "Named Delta", "Named Echo" };
        var options = new PdfOptions {
            CompressContentStreams = false
        };
        foreach (string family in families) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, fontData));
        }

        PdfOptions clone = options.Clone();
        var runs = families
            .Select((family, index) => new TextRun(
                (index == 0 ? string.Empty : " ") + family,
                bold: index == 1,
                italic: index == 2,
                fontFamily: family))
            .ToArray();
        byte[] bytes = PdfDocument.Create(clone)
            .Paragraph(paragraph => paragraph.Runs(runs))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Equal(families.Length, clone.NamedFontFamilies.Count);
        Assert.Equal(families.Length * 2, Regex.Matches(raw, @"/BaseFont /Named(?:Alpha|Bravo|Charlie|Delta|Echo)-(?:Regular|Bold|Italic)").Count);
        Assert.Equal(string.Join(" ", families), text.Trim());
        Assert.DoesNotContain("/BaseFont /Times-Roman", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Courier", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFamily_SnapshotsFaceBytesForReusableRegistration() {
        byte[] regular = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        byte[] bold = { 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0 };
        var family = new PdfEmbeddedFontFamily("Snapshot Family", regular, bold);
        regular[0] = 255;
        bold[5] = 255;

        PdfOptions options = new PdfOptions().UseFontFamily(family);
        byte[] readback = family.Regular;
        readback[0] = 127;

        Assert.Equal("Snapshot Family", family.FamilyName);
        Assert.Equal(0, family.Regular[0]);
        Assert.Equal(1, family.Bold![5]);
        Assert.Equal("Snapshot Family-Regular", options.EmbeddedFonts[PdfStandardFont.Helvetica].FontName);
        Assert.Equal("Snapshot Family-Bold", options.EmbeddedFonts[PdfStandardFont.HelveticaBold].FontName);
        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.FooterFont);
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFontFamily("Broken", Array.Empty<byte>()));
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemLoadsInstalledTrueTypeFamily() {
        if (!TryFindInstalledSystemFontFamily(out PdfEmbeddedFontFamily? family)) {
            return;
        }

        PdfOptions options = new PdfOptions().UseFontFamily(family);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily(family)
            .Paragraph(paragraph => paragraph
                .Text("System regular ")
                .Bold("system bold ")
                .Italic("system italic ")
                .Text("system done"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.NotEmpty(family.Regular);
        Assert.True(family.Bold != null || family.Italic != null || family.BoldItalic != null);
        Assert.Contains(PdfStandardFont.Helvetica, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.HelveticaBold, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.HelveticaOblique, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.HelveticaBoldOblique, options.EmbeddedFonts.Keys);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("System regular system bold system italic system done", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFamily_FromSystemThrowsWhenFamilyIsMissing() {
        Assert.Throws<FileNotFoundException>(() =>
            PdfEmbeddedFontFamily.FromSystem("OfficeIMO Missing Font Family 404"));
    }

    [Fact]
    public void NamedOfficeFontRegistrationHonorsInstalledCandidateOrderBeforeReuse() {
        string[] candidates = {
            "Arial", "Calibri", "Helvetica", "Times New Roman",
            "DejaVu Sans", "Liberation Sans", "Noto Sans"
        };
        PdfEmbeddedFontFamily[] installed = candidates
            .Select(candidate => PdfEmbeddedFontFamily.TryFromSystem(
                candidate,
                out PdfEmbeddedFontFamily? family)
                ? family
                : null)
            .Where(family => family != null)
            .Cast<PdfEmbeddedFontFamily>()
            .GroupBy(family => family.FamilyName, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .Take(2)
            .ToArray();
        if (installed.Length < 2) return;
        var options = new PdfOptions().RegisterNamedFontFamily(installed[1]);

        Assert.True(options.TryRegisterNamedOfficeFontFamily(
            installed[0].FamilyName + ", " + installed[1].FamilyName,
            out string? registeredFamilyName));

        Assert.Equal(installed[0].FamilyName, registeredFamilyName, ignoreCase: true);
    }

    [Fact]
    public void PdfOptions_UseOfficeFontFamilyFallsBackThroughFamilyListToStandardFont() {
        PdfOptions options = new PdfOptions()
            .UseOfficeFontFamily("OfficeIMO Missing Display, Consolas, monospace", embedSystemFont: false);

        Assert.Equal(PdfStandardFont.Courier, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Courier, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Courier, options.FooterFont);
    }

    [Fact]
    public void PdfOptions_UseTextFallbacksPrefersTextCandidateWhenOnlyOneFallbackSlotIsAvailable() {
        if (!DefaultTextSymbolFallbackFontIsAvailable()) {
            return;
        }

        PdfEmbeddedFontFallbackSet? fallbackSet = new PdfOptions()
            .UseTextFallbacks()
            .EmbeddedFontFallbacks;
        if (fallbackSet == null ||
            fallbackSet.Candidates.Count != 1) {
            return;
        }

        Assert.DoesNotContain("Emoji", fallbackSet.Candidates[0].FontName, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PdfOptions_UseTextFallbacksCoversCheckMarkWhenOnlyOneSymbolSlotIsAvailable() {
        if (!PdfEmbeddedFontFamily.TryFromSystem("Segoe UI Symbol", out _) &&
            !PdfEmbeddedFontFamily.TryFromSystem("DejaVu Sans", out _)) {
            return;
        }

        var options = new PdfOptions();
        options.UseTextFallbacks(
            PdfTextFallbackFeatures.SymbolAndEmojiFonts,
            new[] { PdfStandardFont.Helvetica },
            allowSystemFontEmbedding: true);

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacks;
        Assert.NotNull(fallbackSet);
        Assert.Single(fallbackSet!.Candidates);
        Assert.True(fallbackSet.PlanText("\u2713").IsFullyCovered);
    }

    [Fact]
    public void PdfOptions_UseTextFallbacksDoesNotAssignAutomaticFallbacksToTimesSlot() {
        if (!DefaultTextSymbolFallbackFontIsAvailable()) {
            return;
        }

        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.TimesRoman,
            HeaderFont = PdfStandardFont.TimesRoman,
            FooterFont = PdfStandardFont.TimesRoman
        }.UseTextFallbacks(PdfTextFallbackFeatures.SymbolAndEmojiFonts);

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacks;
        if (fallbackSet == null) {
            return;
        }

        Assert.DoesNotContain(
            PdfStandardFont.TimesRoman,
            fallbackSet.FontSlots.Select(PdfStandardFontMapper.GetFontFamily));
    }

    [Fact]
    public void PdfOptions_UseTextFallbacksKeepsSymbolFallbackSlotAheadOfMonospaceFallback() {
        if (!DefaultTextSymbolFallbackFontIsAvailable()) {
            return;
        }

        var options = new PdfOptions()
            .UseOfficeFontFamily(PdfOptions.DefaultDocumentFontFamilyFallback)
            .UseTextFallbacks();

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacks;
        if (fallbackSet == null) {
            return;
        }

        Assert.Contains(
            PdfStandardFont.Courier,
            fallbackSet.FontSlots.Select(PdfStandardFontMapper.GetFontFamily));
        Assert.False(options.TryRegisterDefaultDocumentMonospaceFontFallback());
    }

    [Fact]
    public void PdfOptions_TryUseDefaultDocumentFontFallbackEmbedsUnicodeCapableGeneratedText() {
        var options = new PdfOptions {
            CompressContentStreams = false
        };
        if (!options.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(polish))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(options.HasEmbeddedStandardFontFamily(options.DefaultFont));
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_TryRegisterDefaultDocumentMonospaceFontFallbackEmbedsCourierSlotWhenAvailable() {
        var options = new PdfOptions();
        if (!options.TryRegisterDefaultDocumentMonospaceFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        Assert.True(options.HasEmbeddedStandardFontFamily(PdfStandardFont.Courier));
    }

    [Fact]
    public void PdfOptions_UseTextFallbacksReturnsOptionsForFluentConfiguration() {
        var options = new PdfOptions();

        PdfOptions returned = options.UseTextFallbacks(PdfTextFallbackFeatures.None);

        Assert.Same(options, returned);
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void PdfDocumentAndPageComposeExposeTextFallbackPresetFluently() {
        PdfDocument document = PdfDocument.Create();
        bool visitedPage = false;

        PdfDocument returned = document
            .UseTextFallbacks(PdfTextFallbackFeatures.None)
            .UseEmbeddedFontFallbacksFromSystem("OfficeIMO Missing Font", maxFallbackFonts: 1)
            .Page(page => {
                visitedPage = true;
                Assert.Same(page, page.UseTextFallbacks(PdfTextFallbackFeatures.None));
                Assert.Same(page, page.UseEmbeddedFontFallbacksFromSystem("OfficeIMO Missing Font", maxFallbackFonts: 1));
            });

        Assert.Same(document, returned);
        Assert.True(visitedPage);
    }

    [Fact]
    public void PdfOptions_UseEmbeddedFontFallbacksFromSystemRegistersAvailableFallbackWithoutCallerSlots() {
        if (!TryFindInstalledSystemFontFamily(out PdfEmbeddedFontFamily? family) ||
            family == null) {
            return;
        }

        var options = new PdfOptions();

        PdfOptions returned = options.UseEmbeddedFontFallbacksFromSystem("OfficeIMO Missing Font, " + family.FamilyName, maxFallbackFonts: 1);

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacks;
        Assert.Same(options, returned);
        Assert.NotNull(fallbackSet);
        Assert.Single(fallbackSet!.Candidates);
        Assert.Single(fallbackSet.FontSlots);
        Assert.Equal(family.FamilyName, fallbackSet.Candidates[0].FontName);
        Assert.True(options.HasEmbeddedStandardFontFamily(fallbackSet.FontSlots[0]));
    }

    [Fact]
    public void PdfOptions_UseEmbeddedFontFallbacksFromSystemPreservesExplicitFallbackSet() {
        var explicitFallback = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Explicit Fallback", CreateMinimalOpenTypeCffFont()) },
            new[] { PdfStandardFont.TimesRoman });
        var options = new PdfOptions().RegisterEmbeddedFontFallbacks(explicitFallback);

        Assert.True(options.TryRegisterEmbeddedFontFallbacksFromSystem("Arial, DejaVu Sans", maxFallbackFonts: 1));

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacks;
        Assert.NotNull(fallbackSet);
        Assert.Equal("Explicit Fallback", fallbackSet!.Candidates[0].FontName);
        Assert.Equal(PdfStandardFont.TimesRoman, fallbackSet.FontSlots[0]);
    }

    [Fact]
    public void PdfOptions_CreateRegisteredFontFamilySlotsNormalizesConfiguredAndEmbeddedFamilies() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.HelveticaBold,
            HeaderFont = PdfStandardFont.TimesItalic,
            FooterFont = PdfStandardFont.CourierBoldOblique
        }.RegisterFontFamily(
            PdfStandardFont.TimesRoman,
            new PdfEmbeddedFontFamily("OfficeIMO Slot Test", CreateMinimalOpenTypeCffFont()));

        HashSet<PdfStandardFont> configuredAndEmbedded = options.CreateRegisteredFontFamilySlots(includeDocumentFontSlots: true);
        HashSet<PdfStandardFont> embeddedOnly = options.CreateRegisteredFontFamilySlots(includeDocumentFontSlots: false);

        Assert.Contains(PdfStandardFont.Helvetica, configuredAndEmbedded);
        Assert.Contains(PdfStandardFont.TimesRoman, configuredAndEmbedded);
        Assert.Contains(PdfStandardFont.Courier, configuredAndEmbedded);
        Assert.DoesNotContain(PdfStandardFont.HelveticaBold, configuredAndEmbedded);
        Assert.DoesNotContain(PdfStandardFont.CourierBoldOblique, configuredAndEmbedded);
        Assert.Equal(new[] { PdfStandardFont.TimesRoman }, embeddedOnly.OrderBy(font => font).ToArray());
    }

    [Fact]
    public void PdfOptions_TryAddOfficeFontFamilyKeyTrimsNormalizesAndDeduplicatesFamilies() {
        var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Assert.True(PdfOptions.TryAddOfficeFontFamilyKey("  Aptos Display  ", registeredFamilies, value => value.ToUpperInvariant(), out string trimmedFamilyName));
        Assert.Equal("Aptos Display", trimmedFamilyName);
        Assert.False(PdfOptions.TryAddOfficeFontFamilyKey("aptos display", registeredFamilies, value => value.ToUpperInvariant(), out _));
        Assert.False(PdfOptions.TryAddOfficeFontFamilyKey("   ", registeredFamilies, value => value, out _));
    }

    [Fact]
    public void PdfOptions_TryRegisterMappedOfficeFontFamilyRegistersMappedSlotOnce() {
        var options = new PdfOptions();
        var registeredFontSlots = new HashSet<PdfStandardFont>();

        Assert.True(options.TryRegisterMappedOfficeFontFamily("Times New Roman", registeredFontSlots, embedSystemFont: false, out PdfStandardFont firstSlot));
        Assert.Equal(PdfStandardFont.TimesRoman, firstSlot);
        Assert.Contains(PdfStandardFont.TimesRoman, registeredFontSlots);

        Assert.True(options.TryRegisterMappedOfficeFontFamily("Times", registeredFontSlots, embedSystemFont: false, out PdfStandardFont secondSlot));
        Assert.Equal(PdfStandardFont.TimesRoman, secondSlot);
        Assert.Equal(new[] { PdfStandardFont.TimesRoman }, registeredFontSlots.ToArray());
    }

    [Fact]
    public void PdfOptions_TrySelectAvailableFontFamilySlotUsesMappedFamilyBeforeSharedPreferenceOrder() {
        var registeredFontSlots = new HashSet<PdfStandardFont>();

        Assert.True(PdfOptions.TrySelectAvailableFontFamilySlot("Courier New", registeredFontSlots, out PdfStandardFont mappedSlot));
        Assert.Equal(PdfStandardFont.Courier, mappedSlot);

        registeredFontSlots.Add(PdfStandardFont.Courier);
        Assert.True(PdfOptions.TrySelectAvailableFontFamilySlot("Courier New", registeredFontSlots, out PdfStandardFont fallbackSlot));
        Assert.Equal(PdfStandardFont.TimesRoman, fallbackSlot);

        registeredFontSlots.Add(PdfStandardFont.TimesRoman);
        registeredFontSlots.Add(PdfStandardFont.Helvetica);
        Assert.False(PdfOptions.TrySelectAvailableFontFamilySlot("Aptos", registeredFontSlots, out _));
    }

    [Fact]
    public void PdfOptions_GetAvailableEmbeddedFallbackFontSlotsSkipsDocumentReservedAndEmbeddedFamilies() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.HelveticaBold,
            HeaderFont = PdfStandardFont.Helvetica,
            FooterFont = PdfStandardFont.HelveticaOblique
        }.RegisterFontFamily(
            PdfStandardFont.TimesRoman,
            new PdfEmbeddedFontFamily("OfficeIMO Fallback Slot Test", CreateMinimalOpenTypeCffFont()));

        PdfStandardFont[] slots = options
            .GetAvailableEmbeddedFallbackFontSlots(3, new[] { PdfStandardFont.CourierBold })
            .ToArray();

        Assert.Empty(slots);
    }

    [Fact]
    public void PdfOptions_RegisterFontFamilyEmbedsSemanticSlotWithoutChangingDefaults() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions {
            CompressContentStreams = false
        }.RegisterFontFamily(
            PdfStandardFont.TimesRoman,
            PdfEmbeddedFontFamily.FromFiles("OfficeIMO Semantic Serif", fontPath));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph
                .Font(PdfStandardFont.TimesRoman)
                .Text("Semantic serif Łódź"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.FooterFont);
        Assert.Contains("/BaseFont /OfficeIMOSemanticSerif-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("Semantic serif Łódź", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesMatchesRenamedTrueTypeMetadata() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string renamedPath = Path.Combine(tempDir, "officeimo-renamed-system-face.ttf");
            File.Copy(fontPath, renamedPath);

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                familyName,
                new[] { renamedPath },
                out PdfEmbeddedFontFamily? family);

            Assert.True(found);
            Assert.NotNull(family);
            Assert.NotEmpty(family!.Regular);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesSkipsReadableMetadataMismatchBeforeFilenameFallback() {
        if (!TryFindSingleInstalledRegularFontFace(out _, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string renamedPath = Path.Combine(tempDir, "OfficeIMOFilenameTrap-Regular.ttf");
            File.Copy(fontPath, renamedPath);

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                "OfficeIMO Filename Trap",
                new[] { renamedPath },
                out PdfEmbeddedFontFamily? family);

            Assert.False(found);
            Assert.Null(family);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_MetadataMatchingDoesNotUseFileNameAliases() {
        Assert.False(PdfEmbeddedFontFamily.IsMetadataFamilyNameMatch("Times", "Times New Roman"));
        Assert.False(PdfEmbeddedFontFamily.IsMetadataFamilyNameMatch("Courier", "Courier New"));
        Assert.True(PdfEmbeddedFontFamily.IsMetadataFamilyNameMatch("Times New Roman", "Times New Roman"));
    }

    [Fact]
    public void PdfEmbeddedFontFamily_MetadataStyleScoringPrefersExactBoldItalic() {
        Assert.True(
            PdfEmbeddedFontFamily.GetMetadataStyleScore("Bold Italic") >
            PdfEmbeddedFontFamily.GetMetadataStyleScore("SemiBold Italic"));
        Assert.True(
            PdfEmbeddedFontFamily.GetMetadataStyleScore("Bold") >
            PdfEmbeddedFontFamily.GetMetadataStyleScore("SemiBold"));
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesMatchesTrueTypeCollectionFace() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string collectionPath = Path.Combine(tempDir, "officeimo-system-face.ttc");
            File.WriteAllBytes(collectionPath, CreateSingleFaceTrueTypeCollection(File.ReadAllBytes(fontPath)));

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                familyName,
                new[] { collectionPath },
                out PdfEmbeddedFontFamily? family);

            Assert.True(found);
            Assert.NotNull(family);
            Assert.NotEmpty(family!.Regular);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesStopsAfterScanBudget() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string[] missingFiles = Enumerable
            .Range(0, PdfEmbeddedFontFamily.MaxSystemFontFilesToInspect)
            .Select(index => Path.Combine(Path.GetTempPath(), "OfficeIMO.Missing." + index.ToString(CultureInfo.InvariantCulture) + ".ttf"))
            .Concat(new[] { fontPath })
            .ToArray();

        bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
            familyName,
            missingFiles,
            out PdfEmbeddedFontFamily? family);

        Assert.False(found);
        Assert.Null(family);
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesSkipsOversizedFontFiles() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string oversizedPath = Path.Combine(tempDir, "OfficeIMO-Oversized.ttf");
            using (FileStream stream = File.Create(oversizedPath)) {
                stream.SetLength(PdfEmbeddedFontFamily.MaxSystemFontFileBytes + 1);
            }

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                familyName,
                new[] { oversizedPath, fontPath },
                out PdfEmbeddedFontFamily? family);

            Assert.True(found);
            Assert.NotNull(family);
            Assert.NotEmpty(family!.Regular);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesRejectsOversizedTrueTypeCollectionFaceCounts() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string collectionPath = Path.Combine(tempDir, "officeimo-oversized-system-face.ttc");
            File.WriteAllBytes(
                collectionPath,
                CreateRepeatedFaceTrueTypeCollection(
                    File.ReadAllBytes(fontPath),
                    PdfEmbeddedFontFamily.MaxTrueTypeCollectionFontsToInspect + 1));

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                familyName,
                new[] { collectionPath },
                out PdfEmbeddedFontFamily? family);

            Assert.False(found);
            Assert.Null(family);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesSkipsMalformedTrueTypeFiles() {
        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string malformedPath = Path.Combine(tempDir, "OfficeIMOMalformed-Regular.ttf");
            File.WriteAllBytes(malformedPath, new byte[] { 0, 1, 0, 0, 0, 255, 255, 255 });

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                "OfficeIMO Malformed",
                new[] { malformedPath },
                out PdfEmbeddedFontFamily? family);

            Assert.False(found);
            Assert.Null(family);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_GetSystemFontRootsIncludesWindowsPerUserFonts() {
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData)) {
            return;
        }

        string expected = Path.Combine(localAppData, "Microsoft", "Windows", "Fonts");
        Assert.Contains(
            PdfEmbeddedFontFamily.GetSystemFontRoots(),
            root => string.Equals(root, expected, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PdfDocument_UseFontFamilyObjectReusesTrueTypeFamilyForGeneratedText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("OfficeIMO Object Font", fontData);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily(family)
            .Header(header => header.Text("Object font header"))
            .Paragraph(paragraph => paragraph
                .Text("Object regular ")
                .Bold("object bold ")
                .Italic("object italic"))
            .Footer(footer => footer.Text("Object font footer {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType2", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CIDToGIDMap /Identity", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Italic", raw, StringComparison.Ordinal);
        Assert.All(ExtractLength1Values(raw), length => Assert.InRange(length, 1, fontData.Length - 1));
        Assert.Contains("Object regular object bold object italic", text, StringComparison.Ordinal);
        Assert.Contains("Object font header", text, StringComparison.Ordinal);
        Assert.Contains("Object font footer 1/1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void EmbeddedFontSubsetting_DoesNotCarryGlyphUsageAcrossWriterCallsWithSameOptions() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var options = new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Reset Font");
        Assert.True(options.TryGetEmbeddedStandardFontProgramForGeneration(PdfStandardFont.Helvetica, out _, out PdfTrueTypeFontProgram? fontProgram));
        Assert.NotNull(fontProgram);
        Assert.True(fontProgram!.TryGetGlyphId('Ł', out int lStrokeGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('ó', out int oAcuteGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('ź', out int zAcuteGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('A', out int aGlyphId));

        byte[] first = PdfWriter.Write(
            PdfDocument.Create(),
            new IPdfBlock[] { new RichParagraphBlock(new[] { TextRun.Normal("Łódź") }, PdfAlign.Left, defaultColor: null) },
            options,
            title: null,
            author: null,
            subject: null,
            keywords: null);
        Assert.NotEmpty(first);
        IReadOnlyList<int> firstUsedGlyphIds = fontProgram.GetUsedGlyphIds();
        Assert.Contains(lStrokeGlyphId, firstUsedGlyphIds);
        Assert.Contains(oAcuteGlyphId, firstUsedGlyphIds);
        Assert.Contains(zAcuteGlyphId, firstUsedGlyphIds);

        byte[] second = PdfWriter.Write(
            PdfDocument.Create(),
            new IPdfBlock[] { new RichParagraphBlock(new[] { TextRun.Normal("A") }, PdfAlign.Left, defaultColor: null) },
            options,
            title: null,
            author: null,
            subject: null,
            keywords: null);

        Assert.NotEmpty(second);
        IReadOnlyList<int> secondUsedGlyphIds = fontProgram.GetUsedGlyphIds();

        Assert.Contains(aGlyphId, secondUsedGlyphIds);
        Assert.DoesNotContain(lStrokeGlyphId, secondUsedGlyphIds);
        Assert.DoesNotContain(oAcuteGlyphId, secondUsedGlyphIds);
        Assert.DoesNotContain(zAcuteGlyphId, secondUsedGlyphIds);
    }

    [Fact]
    public void PdfTrueTypeFontProgram_ShapeTextProducesStableUnicodeGlyphRunForMultilingualText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        string text = "Zażółć gęślą jaźń Ελλάδα Київ";
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(File.ReadAllBytes(fontPath), "OfficeIMO Shape Font");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        PdfGlyphRun first = fontProgram.ShapeText(text);
        fontProgram.ResetGlyphUsage();
        PdfGlyphRun second = fontProgram.ShapeText(text);

        Assert.False(first.HasMissingGlyphs);
        Assert.Empty(first.Diagnostics);
        Assert.Equal(CountUnicodeScalars(text), first.Glyphs.Count);
        Assert.Equal(first.TotalAdvanceWidth1000, second.TotalAdvanceWidth1000);
        Assert.Equal(
            first.Glyphs.Select(glyph => glyph.GlyphId).ToArray(),
            second.Glyphs.Select(glyph => glyph.GlyphId).ToArray());
        Assert.Contains(first.Glyphs, glyph => glyph.UnicodeScalar == 'ż');
        Assert.Contains(first.Glyphs, glyph => glyph.UnicodeScalar == 'Ε');
        Assert.Contains(first.Glyphs, glyph => glyph.UnicodeScalar == 'К');
        Assert.Equal(first.ToGlyphHex(), fontProgram.EncodeTextAsGlyphHex(text));
        Assert.Equal(first.TotalAdvanceWidth1000 * 12D / 1000D, fontProgram.MeasureTextWidth(text, 12D), precision: 6);
    }

    [Fact]
    public void PdfTrueTypeFontProgram_ShapeTextCanPreflightMissingGlyphsWithoutThrowing() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        string unsupportedScalar = char.ConvertFromUtf32(0x10FFFF);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(File.ReadAllBytes(fontPath), "OfficeIMO Shape Font");

        PdfGlyphRun glyphRun = fontProgram.ShapeText(
            unsupportedScalar,
            PdfTextShapingOptions.ForDiagnostics("shape-preflight", fontProgram.FontName));

        Assert.Empty(glyphRun.Glyphs);
        Assert.True(glyphRun.HasMissingGlyphs);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(glyphRun.Diagnostics);
        Assert.Equal("missing-embedded-font-glyph", diagnostic.Code);
        Assert.Equal("shape-preflight", diagnostic.Source);
        Assert.Equal("U+10FFFF", diagnostic.CodePoint);

        ArgumentException exception = Assert.Throws<ArgumentException>(() => fontProgram.ShapeText(unsupportedScalar));
        Assert.Contains("not covered by the embedded TrueType font", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_UseFontFamilyScopesFamilyToComposedPageHeaderFooterAndBody() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var family = PdfEmbeddedFontFamily.FromFiles("OfficeIMO Page Font", fontPath);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Before page font"))
            .Page(page => page
                .UseFontFamily(family)
                .Header(header => header.Text("Page font header"))
                .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Page font body"))))
                .Footer(footer => footer.Text("Page font footer {page}/{pages}")))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/BaseFont /OfficeIMOPageFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("Before page font", text, StringComparison.Ordinal);
        Assert.Contains("Page font header", text, StringComparison.Ordinal);
        Assert.Contains("Page font body", text, StringComparison.Ordinal);
        Assert.Contains("Page font footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_UseFontFamilyScopesFamilyToTextFieldAppearance() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var family = PdfEmbeddedFontFamily.FromFiles("OfficeIMO Page Form Font", fontPath);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Page(page => page
                .UseFontFamily(family)
                .Content(content => content.Item(item => item.TextField("Scoped.Name", value: "Lodz"))))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int embeddedFontObjectId = FindObjectIdBefore(raw, "/Subtype /Type0");

        Assert.Contains("/BaseFont /OfficeIMOPageFormFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/Resources << /Font << /Helv " + embeddedFontObjectId.ToString(CultureInfo.InvariantCulture) + " 0 R >>", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void DefaultTextStyle_FontFamilyDoesNotRewriteExistingHeaderOrFooterFonts() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("OfficeIMO Style Font", fontData);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Header(header => header.Font(PdfStandardFont.TimesRoman).Text("Times header"))
            .Footer(footer => footer.Font(PdfStandardFont.Courier).Text("Courier footer"))
            .DefaultTextStyle(style => style.FontFamily(family).FontSize(12))
            .Paragraph(paragraph => paragraph.Text("Styled body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/BaseFont /OfficeIMOStyleFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Times-Roman", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Courier", raw, StringComparison.Ordinal);
        Assert.Contains("Times header", text, StringComparison.Ordinal);
        Assert.Contains("Styled body", text, StringComparison.Ordinal);
        Assert.Contains("Courier footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_UseFontFamilyWritesUnicodeGlyphsAndToUnicodeExtraction() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Unicode Font", fontPath)
            .Header(header => header.Text("Nagłówek Łódź"))
            .Paragraph(paragraph => paragraph
                .Text(polish + " ")
                .Bold("Śląsk ")
                .Italic("źrebak"))
            .Footer(footer => footer.Text("Stopka Łódź {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-Identity-Glyph-UCS", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
        Assert.Contains("Nagłówek Łódź", text, StringComparison.Ordinal);
        Assert.Contains("Śląsk źrebak", text, StringComparison.Ordinal);
        Assert.Contains("Stopka Łódź 1/1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextDiagnostics_UsesEmbeddedFontCoverageForGeneratedText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .UseFontFamily("OfficeIMO Diagnostic Font", fontPath);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(
            "Zażółć gęślą jaźń",
            options,
            PdfStandardFont.Helvetica,
            "body");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void PdfTextDiagnostics_ReportsEmbeddedFontMissingGlyphAsConversionWarning() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .UseFontFamily("OfficeIMO Diagnostic Font", fontPath);
        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(
            "Snowman \u2603",
            options,
            PdfStandardFont.Helvetica,
            "body",
            "PdfParagraph[0]");

        if (diagnostics.Count == 0) {
            return;
        }

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);
        PdfConversionWarning warning = diagnostic.ToConversionWarning("OfficeIMO.Tests");

        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
        Assert.Equal("U+2603", diagnostic.CodePoint);
        Assert.Contains("embedded TrueType font", diagnostic.Encoding, StringComparison.Ordinal);
        Assert.Contains("fallback", diagnostic.Remediation, StringComparison.Ordinal);
        Assert.Equal("PdfParagraph[0]", diagnostic.Location);
        Assert.Equal(diagnostic.Encoding, warning.Details["encoding"]);
        Assert.Equal(diagnostic.Remediation, warning.Details["remediation"]);
        Assert.Equal(diagnostic.Location, warning.Details["location"]);
    }

    [Fact]
    public void PdfDocument_EmbeddedFontMissingGlyphThrowsDiagnosticBackedException() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        string missingGlyph = char.ConvertFromUtf32(0x10FFFF);
        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() =>
            PdfDocument.Create()
                .UseFontFamily("OfficeIMO Diagnostic Font", fontPath)
                .Paragraph(paragraph => paragraph.Text("Missing " + missingGlyph))
                .ToBytes());

        Assert.Contains("U+10FFFF", exception.Message, StringComparison.Ordinal);
        Assert.Contains("embedded TrueType font", exception.Message, StringComparison.Ordinal);
        Assert.Contains("fallback", exception.Message, StringComparison.Ordinal);
        Assert.Equal("unsupported-text-glyph", exception.Data["code"]);
        Assert.Equal("PdfParagraph", exception.Data["source"]);
        Assert.Equal("PdfParagraph[0].Run[0]", exception.Data["location"]);
        Assert.Equal(0, exception.Data["runIndex"]);
        Assert.Equal("U+10FFFF", exception.Data["codePoint"]);
        Assert.Equal(1, exception.Data["diagnosticsCount"]);
        Assert.Contains("embedded TrueType font", (string)exception.Data["encoding"]!, StringComparison.Ordinal);
        Assert.Contains("fallback", (string)exception.Data["remediation"]!, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextDiagnostics_UsesStyledEmbeddedFontSlotsForRuns() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .UseFontFamily("OfficeIMO Diagnostic Family", fontPath);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
            new[] {
                TextRun.Bolded("Bold Łódź"),
                TextRun.Italicized("Italic Łódź"),
                TextRun.BoldItalic("Bold italic Łódź")
            },
            options,
            PdfStandardFont.Helvetica,
            "body");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void PdfDocument_UseFontFamilySubsetsUnicodeFontDeterministically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        byte[] first = CreateMultilingualBusinessReport(fontPath);
        byte[] second = CreateMultilingualBusinessReport(fontPath);
        string raw = Encoding.ASCII.GetString(first);
        string text = PdfReadDocument.Open(first).ExtractText();

        Assert.Equal(first, second);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.All(ExtractLength1Values(raw), length => Assert.InRange(length, 1, fontData.Length - 1));
        Assert.Contains("Q2 multilingual revenue report", text, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", text, StringComparison.Ordinal);
        Assert.Contains("Ελλάδα", text, StringComparison.Ordinal);
        Assert.Contains("Київ", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfConversionReport_AddTextDiagnosticsSurfacesSharedWarnings() {
        var report = new PdfConversionReport();

        report.AddTextDiagnostics(
            PdfTextDiagnostics.AnalyzeWinAnsiText("Status Ω", "word:paragraph[1]"),
            "OfficeIMO.Word.Pdf");

        PdfConversionWarning warning = Assert.Single(report.Warnings);
        Assert.True(report.HasWarnings);
        Assert.Equal("OfficeIMO.Word.Pdf", warning.Converter);
        Assert.Equal("unsupported-text-glyph", warning.Code);
        Assert.Equal("word:paragraph[1]", warning.Source);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
    }

    [Fact]
    public void PdfDocumentConversionResult_SnapshotsConversionReport() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "sample-warning",
            "source[1]",
            "Sample warning."));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Conversion result")),
            report);

        report.Clear();

        PdfConversionWarning warning = Assert.Single(result.Warnings);
        Assert.True(result.HasWarnings);
        Assert.False(report.HasWarnings);
        Assert.Equal("sample-warning", warning.Code);
        Assert.Contains("Conversion result", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);

        PdfDocumentConversionResult processed = result.Process(document => document.UpdateMetadata(title: "Processed conversion result"));

        PdfConversionWarning processedWarning = Assert.Single(processed.Warnings);
        Assert.Equal("sample-warning", processedWarning.Code);
        Assert.Equal("Processed conversion result", processed.Value.Inspect().Metadata.Title);
        Assert.Contains("Conversion result", PdfReadDocument.Open(processed.ToBytes()).ExtractText(), StringComparison.Ordinal);

        using var output = new MemoryStream();
        PdfSaveResult saveResult = processed.TrySave(output);

        Assert.True(saveResult.Succeeded);
        Assert.True(saveResult.BytesWritten > 0);
        Assert.Equal("sample-warning", Assert.Single(saveResult.Warnings).Code);
        Assert.True(saveResult.HasWarnings);
        Assert.True(output.Length > 0);
        PdfSaveResult throwingSaveResult = processed.Save(new MemoryStream());
        Assert.True(throwingSaveResult.Succeeded);
        Assert.Equal("sample-warning", Assert.Single(throwingSaveResult.Warnings).Code);
        Assert.True(throwingSaveResult.Pipeline.Succeeded);
    }

    [Fact]
    public void PdfDocumentConversionResult_RefreshesWarningsEmittedAfterResultCreation() {
        var report = new PdfConversionReport();
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Save diagnostics result")),
            report);

        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "save-time-warning",
            "writer",
            "Warning emitted while serializing the PDF."));

        Assert.False(result.HasWarnings);

        byte[] bytes = result.ToBytes();

        Assert.NotEmpty(bytes);
        PdfConversionWarning warning = Assert.Single(result.Warnings);
        Assert.Equal("save-time-warning", warning.Code);
    }

    [Fact]
    public async System.Threading.Tasks.Task PdfDocumentConversionResult_AsyncSavePreservesConversionReportSnapshot() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "async-sample-warning",
            "source[async]",
            "Async sample warning."));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Async conversion result")),
            report);

        report.Clear();

        using var stream = new MemoryStream();
        PdfSaveResult streamResult = await result.TrySaveAsync(stream);

        Assert.True(streamResult.Succeeded);
        Assert.True(streamResult.BytesWritten > 0);
        Assert.Equal("async-sample-warning", Assert.Single(streamResult.Warnings).Code);
        Assert.True(stream.Length > 0);
        Assert.True(result.HasWarnings);
        Assert.Equal("async-sample-warning", Assert.Single(result.Warnings).Code);

        using var chainedStream = new MemoryStream();
        PdfSaveResult asyncSaveResult = await result.SaveAsync(chainedStream);
        Assert.True(asyncSaveResult.Succeeded);
        Assert.Equal("async-sample-warning", Assert.Single(asyncSaveResult.Warnings).Code);
        Assert.True(asyncSaveResult.Pipeline.Succeeded);
        Assert.True(chainedStream.Length > 0);

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.ConversionResult.Async", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string tryPath = Path.Combine(directory, "try-save.pdf");
            string savePath = Path.Combine(directory, "save.pdf");

            PdfSaveResult pathResult = await result.TrySaveAsync(tryPath);
            PdfSaveResult pathSaveResult = await result.SaveAsync(savePath);

            Assert.True(pathResult.Succeeded);
            Assert.Equal("async-sample-warning", Assert.Single(pathResult.Warnings).Code);
            Assert.True(File.Exists(tryPath));
            Assert.True(new FileInfo(tryPath).Length > 0);
            Assert.True(pathSaveResult.Succeeded);
            Assert.True(File.Exists(savePath));
            Assert.True(new FileInfo(savePath).Length > 0);
            Assert.Equal("async-sample-warning", Assert.Single(pathSaveResult.Warnings).Code);
            Assert.True(pathSaveResult.Pipeline.Succeeded);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeAdvancedTextLayoutReportsComplexScriptRequirements() {
        IReadOnlyList<PdfTextShapingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeAdvancedTextLayout(
            "\u0645\u0631\u062D\u0628\u0627 office",
            "word:paragraph[rtl]");

        PdfTextShapingDiagnostic bidi = Assert.Single(diagnostics, item => item.Code == "unsupported-bidirectional-text-layout");
        PdfTextShapingDiagnostic shaping = Assert.Single(diagnostics, item => item.Code == "unsupported-complex-script-shaping");
        Assert.Equal("word:paragraph[rtl]", bidi.Source);
        Assert.Equal("right-to-left", bidi.Script);
        Assert.Equal("Arabic", shaping.Script);
        Assert.Equal("U+0645", bidi.CodePoint);

        var report = new PdfConversionReport();
        report.AddTextShapingDiagnostics(diagnostics, "OfficeIMO.Word.Pdf");
        PdfConversionWarning warning = Assert.Single(report.Warnings, item => item.Code == "unsupported-complex-script-shaping");
        Assert.Equal(PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal(PdfLayoutDiagnosticKind.SimplifiedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("Arabic", warning.Details["script"]);
    }

    [Fact]
    public void PdfOptions_ReportDiagnosticsToSurfacesAdvancedTextLayoutWarningsBeforeWinAnsiFailure() {
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Text("\u0645\u0631\u062D\u0628\u0627"))
                .ToBytes());

        Assert.Contains("WinAnsiEncoding", exception.Message, StringComparison.Ordinal);
        PdfConversionWarning shapingWarning = Assert.Single(report.Warnings, item => item.Code == "unsupported-complex-script-shaping");
        PdfConversionWarning bidiWarning = Assert.Single(report.Warnings, item => item.Code == "unsupported-bidirectional-text-layout");
        Assert.Equal(PdfConversionWarningSeverity.Warning, shapingWarning.Severity);
        Assert.Equal(PdfLayoutDiagnosticKind.SimplifiedContent, bidiWarning.LayoutDiagnostic!.Kind);
        Assert.Contains(report.Warnings, item => item.Code == "unsupported-text-glyph" && item.Severity == PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void PdfOptions_ReportDiagnosticsToResetsTextDiagnosticDeduplicationForNewReport() {
        var first = new PdfConversionReport();
        var second = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(first, "OfficeIMO.Tests");

        Assert.ThrowsAny<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Text("\u0645\u0631\u062D\u0628\u0627"))
                .ToBytes());

        options.ReportDiagnosticsTo(second, "OfficeIMO.Tests");

        Assert.ThrowsAny<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Text("\u0645\u0631\u062D\u0628\u0627"))
                .ToBytes());

        Assert.Contains(first.Warnings, item => item.Code == "unsupported-complex-script-shaping");
        Assert.Contains(second.Warnings, item => item.Code == "unsupported-complex-script-shaping");
    }

    [Fact]
    public void PdfFontDiagnostics_AnalyzeEmbeddedFontReportsOpenTypeCffBeforeRendering() {
        IReadOnlyList<PdfFontEmbeddingDiagnostic> diagnostics = PdfFontDiagnostics.AnalyzeEmbeddedFont(
            CreateMinimalOpenTypeCffFont(),
            "word:styles[Body]",
            "OfficeIMO CFF");

        PdfFontEmbeddingDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal("unsupported-opentype-cff-font", diagnostic.Code);
        Assert.Equal("word:styles[Body]", diagnostic.Source);
        Assert.Equal("OfficeIMO CFF", diagnostic.FontName);
        Assert.Equal("OpenType/CFF", diagnostic.Format);
        Assert.Contains("could not be parsed", diagnostic.Message, StringComparison.Ordinal);

        var report = new PdfConversionReport();
        report.AddFontDiagnostics(diagnostics, "OfficeIMO.Word.Pdf");
        PdfConversionWarning warning = Assert.Single(report.Warnings);
        Assert.Equal("OfficeIMO.Word.Pdf", warning.Converter);
        Assert.Equal("unsupported-opentype-cff-font", warning.Code);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("OpenType/CFF", warning.Details["format"]);
        Assert.Equal("OfficeIMO CFF", warning.Details["fontName"]);
    }

    [Fact]
    public void PdfOpenTypeFontInspector_InspectsRealOpenTypeCffFontCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        PdfOpenTypeFontInfo info = PdfOpenTypeFontInspector.Inspect(
            File.ReadAllBytes(fontPath!),
            "OfficeIMO Source Serif CFF");

        Assert.True(info.IsOpenTypeCff);
        Assert.False(info.IsTrueType);
        Assert.Equal("OTTO", info.ScalerType);
        Assert.True(info.GlyphCount > 1000);
        Assert.True(info.CffTableLength > 1000);
        Assert.True(info.UnicodeScalarCount > 500);
        Assert.True(info.HasGlyphSubstitutionTable);
        Assert.True(info.HasGlyphPositioningTable);
        Assert.Contains("liga", info.GlyphSubstitutionFeatureTags);
        Assert.Contains("mark", info.GlyphPositioningFeatureTags);
        Assert.Contains("mkmk", info.GlyphPositioningFeatureTags);
        Assert.True(info.ContainsUnicodeScalar('A'));
        Assert.True(info.ContainsUnicodeScalar('Ł'));
        Assert.False(info.ContainsUnicodeScalar(0x10FFFF));
        Assert.Equal("OfficeIMOSourceSerifCFF", info.FontName);
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeAdvancedTextLayoutWithFontReportsOpenTypeFeatureGaps() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        IReadOnlyList<PdfTextShapingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeAdvancedTextLayout(
            "office cafe\u0301",
            File.ReadAllBytes(fontPath!),
            "word:paragraph[features]",
            "OfficeIMO Source Serif CFF");

        PdfTextShapingDiagnostic ligature = Assert.Single(diagnostics, item => item.Code == "unsupported-font-ligature-substitution");
        PdfTextShapingDiagnostic mark = Assert.Single(diagnostics, item => item.Code == "unsupported-font-mark-positioning");

        Assert.Equal("word:paragraph[features]", ligature.Source);
        Assert.Equal("OpenType GSUB ligature", ligature.Script);
        Assert.Equal("U+0066", ligature.CodePoint);
        Assert.Contains("OfficeIMOSourceSerifCFF", ligature.Message, StringComparison.Ordinal);
        Assert.Equal("OpenType GPOS mark", mark.Script);
        Assert.Equal("U+0301", mark.CodePoint);

        var report = new PdfConversionReport();
        report.AddTextShapingDiagnostics(diagnostics, "OfficeIMO.Word.Pdf");

        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-mark-positioning");
    }

    [Fact]
    public void PdfOptions_ReportDiagnosticsToSurfacesOpenTypeFeatureWarningsForEmbeddedFont() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "OfficeIMO Source Serif CFF");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("office"))
            .ToBytes();

        Assert.NotEmpty(bytes);
        PdfConversionWarning warning = Assert.Single(report.Warnings, item => item.Code == "unsupported-font-ligature-substitution");
        Assert.Equal("OfficeIMO.Tests", warning.Converter);
        Assert.Equal("OpenType GSUB ligature", warning.Details["script"]);
    }

    [Fact]
    public void PdfFontDiagnostics_AnalyzeEmbeddedFontAcceptsParseableOpenTypeCffFont() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        IReadOnlyList<PdfFontEmbeddingDiagnostic> diagnostics = PdfFontDiagnostics.AnalyzeEmbeddedFont(
            File.ReadAllBytes(fontPath!),
            "word:styles[Body]",
            "OfficeIMO Source Serif CFF");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void PdfFontDiagnostics_AnalyzeEmbeddedFontRejectsOpenTypeCffFontsThatEmbeddingParserRejects() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] restrictedFont = CreateEmbeddingRestrictedOpenTypeCffFont(File.ReadAllBytes(fontPath!));

        PdfOpenTypeFontInfo info = PdfOpenTypeFontInspector.Inspect(restrictedFont, "OfficeIMO Restricted CFF");
        IReadOnlyList<PdfFontEmbeddingDiagnostic> diagnostics = PdfFontDiagnostics.AnalyzeEmbeddedFont(
            restrictedFont,
            "word:styles[Restricted]",
            "OfficeIMO Restricted CFF");

        Assert.True(info.IsOpenTypeCff);
        PdfFontEmbeddingDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal("unsupported-opentype-cff-font", diagnostic.Code);
        Assert.Equal("OpenType/CFF", diagnostic.Format);
        Assert.Contains("fsType", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOpenTypeCffFontProgram_ParsesMetricsAndEncodesGlyphIdsFromRealFont() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        PdfOpenTypeCffFontProgram program = PdfOpenTypeCffFontProgram.Parse(
            File.ReadAllBytes(fontPath!),
            "OfficeIMO Source Serif CFF");

        Assert.Equal("OfficeIMOSourceSerifCFF", program.FontName);
        Assert.True(program.GlyphCount > 1000);
        Assert.True(program.CffTableLength > 1000);
        Assert.True(program.UnitsPerEm > 0);
        Assert.Equal(4, program.FontBBox.Length);
        Assert.True(program.Ascent > 0);
        Assert.True(program.Descent < 0);
        Assert.True(program.CapHeight > 0);
        Assert.True(program.TryGetGlyphId('Ł', out int polishGlyphId));
        Assert.True(program.GetGlyphWidth1000(polishGlyphId) > 0);
        Assert.True(program.MeasureTextWidth("AŁéΩ", 12) > 0);

        string glyphHex = program.EncodeTextAsGlyphHex("AŁéΩ");

        Assert.Equal(16, glyphHex.Length);
        Assert.DoesNotContain("0000", glyphHex, StringComparison.Ordinal);
        Assert.Contains(polishGlyphId, program.GetUsedGlyphIds());
    }

    [Fact]
    public void PdfOpenTypeCffFontProgram_BuildsFontFile3DictionaryAndToUnicodeBoundary() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        PdfOpenTypeCffFontProgram program = PdfOpenTypeCffFontProgram.Parse(
            File.ReadAllBytes(fontPath!),
            "OfficeIMO Source Serif CFF");
        Assert.True(program.TryGetGlyphId('Ł', out int polishGlyphId));
        _ = program.EncodeTextAsGlyphHex("AŁ");

        string descriptor = PdfStandardFontDictionaryBuilder.BuildOpenTypeCffFontDescriptorObject(program, 10);
        string descendant = PdfStandardFontDictionaryBuilder.BuildCidFontType0DescendantObject(program, 11);
        string type0 = PdfStandardFontDictionaryBuilder.BuildEmbeddedType0FontObject(program, 12, 13);
        string toUnicode = Encoding.ASCII.GetString(PdfToUnicodeCMapBuilder.BuildIdentityGlyphToUnicodeCMap(program));

        Assert.Contains("/FontFile3 10 0 R", descriptor, StringComparison.Ordinal);
        Assert.DoesNotContain("/FontFile2", descriptor, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType0", descendant, StringComparison.Ordinal);
        Assert.Contains("/W [", descendant, StringComparison.Ordinal);
        Assert.DoesNotContain("/CIDToGIDMap", descendant, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", type0, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", type0, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode 13 0 R", type0, StringComparison.Ordinal);
        Assert.Contains("<" + polishGlyphId.ToString("X4", CultureInfo.InvariantCulture) + "> <0141>", toUnicode, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_EmbedStandardFontWritesOpenTypeCffFontFile3Output() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "OfficeIMO Source Serif CFF");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("CFF Łódź A"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.StartsWith("%PDF-1.6", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/FontFile2", raw, StringComparison.Ordinal);
        Assert.Contains("CFF Łódź A", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "opentype-cff-font-output-not-enabled");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-opentype-cff-font");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "opentype-cff-charstrings-not-subset");
    }

    [Fact]
    public void PdfOptions_EmbedStandardFontReplacesCachedOpenTypeCffProgram() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontBytes = File.ReadAllBytes(fontPath!);
        var options = new PdfOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO First CFF");

        Assert.True(options.TryGetEmbeddedStandardOpenTypeCffFontProgram(PdfStandardFont.Helvetica, out PdfOpenTypeCffFontProgram? firstProgram));
        Assert.NotNull(firstProgram);
        Assert.Equal("OfficeIMOFirstCFF", firstProgram!.FontName);

        options.EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO Second CFF");

        Assert.True(options.TryGetEmbeddedStandardOpenTypeCffFontProgram(PdfStandardFont.Helvetica, out PdfOpenTypeCffFontProgram? secondProgram));
        Assert.NotNull(secondProgram);
        Assert.Equal("OfficeIMOSecondCFF", secondProgram!.FontName);

        options.ClearEmbeddedStandardFonts();

        Assert.False(options.TryGetEmbeddedStandardOpenTypeCffFontProgram(PdfStandardFont.Helvetica, out _));
    }

    [Fact]
    public void PdfOptions_ClearEmbeddedStandardFontsClearsFallbackPlans() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Source Serif CFF", File.ReadAllBytes(fontPath!)) },
            new[] { PdfStandardFont.TimesRoman });
        var options = new PdfOptions {
            CompressContentStreams = false,
            CompressEmbeddedFonts = false
        }.RegisterEmbeddedFontFallbacks(fallbackSet);

        options.ClearEmbeddedStandardFonts();

        Assert.Null(options.EmbeddedFontFallbacks);
        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Text("Fallback cleared Łódź"))
                .ToBytes());
        Assert.Contains("WinAnsiEncoding", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_RendersOpenTypeCffCandidateThroughFluentFallbackText() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Source Serif CFF", File.ReadAllBytes(fontPath!)) },
            new[] { PdfStandardFont.TimesRoman });
        PdfTextFallbackPlan plan = fallbackSet.PlanText("CFF fallback Łódź", "pdf:paragraph[1]");

        Assert.True(plan.IsFullyCovered);
        Assert.Empty(plan.Diagnostics);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Paragraph(paragraph => paragraph.FallbackText(fallbackSet, "CFF fallback Łódź", "pdf:paragraph[1]"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.StartsWith("%PDF-1.6", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/FontFile2", raw, StringComparison.Ordinal);
        Assert.Contains("CFF fallback Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_EmbeddedFontFallbacksPropertyRegistersFallbackFontSlots() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Source Serif CFF", File.ReadAllBytes(fontPath!)) },
            new[] { PdfStandardFont.TimesRoman });
        var options = new PdfOptions {
            CompressContentStreams = false,
            CompressEmbeddedFonts = false,
            EmbeddedFontFallbacks = fallbackSet
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("CFF property fallback Łódź"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(PdfStandardFont.TimesRoman, options.EmbeddedFonts.Keys);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("CFF property fallback Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_ReportDiagnosticsToSurfacesMalformedOpenTypeCffFontBeforeThrowing() {
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateMinimalOpenTypeCffFont(), "OfficeIMO CFF");

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Text("Unsupported configured font"))
                .ToBytes());

        Assert.Contains("OpenType", exception.Message, StringComparison.Ordinal);
        PdfConversionWarning warning = Assert.Single(report.Warnings, item => item.Code == "unsupported-opentype-cff-font");
        Assert.Equal("OfficeIMO.Tests", warning.Converter);
        Assert.Equal("embedded-font:Helvetica", warning.Source);
        Assert.Equal("OpenType/CFF", warning.Details["format"]);
        Assert.Equal("OfficeIMO CFF", warning.Details["fontName"]);
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeEmbeddedFontTextReportsMissingGlyphsBeforeRendering() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        string text = "Invoice " + char.ConvertFromUtf32(0x10FFFF);

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(
            PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontData, "word:paragraph[2]", "OfficeIMO Coverage Font"));

        Assert.Equal("missing-embedded-font-glyph", diagnostic.Code);
        Assert.Equal("U+10FFFF", diagnostic.CodePoint);
        Assert.Equal("word:paragraph[2]", diagnostic.Source);
        Assert.Contains("OfficeIMO Coverage Font", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeEmbeddedFontTextAcceptsCoveredGlyphs() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(
            "Invoice Cafe 123",
            fontData,
            "word:paragraph[3]",
            "OfficeIMO Coverage Font");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeEmbeddedFontTextSupportsOpenTypeCffBytes() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontData = File.ReadAllBytes(fontPath!);
        string text = "CFF Łódź " + char.ConvertFromUtf32(0x10FFFF);

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(
            PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontData, "word:paragraph[4]", "OfficeIMO CFF Coverage Font"));

        Assert.Equal("missing-embedded-font-glyph", diagnostic.Code);
        Assert.Equal("U+10FFFF", diagnostic.CodePoint);
        Assert.Equal("word:paragraph[4]", diagnostic.Source);
        Assert.Contains("OpenType/CFF", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("OfficeIMO CFF Coverage Font", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeEmbeddedFontTextRunsSupportsOpenTypeCffBytes() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontData = File.ReadAllBytes(fontPath!);
        var runs = new[] {
            new TextRun("CFF Łódź"),
            new TextRun("\n"),
            new TextRun(" tail " + char.ConvertFromUtf32(0x10FFFF))
        };

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(
            PdfTextDiagnostics.AnalyzeEmbeddedFontTextRuns(runs, fontData, "word:runs[4]", "OfficeIMO CFF Coverage Font"));

        Assert.Equal("missing-embedded-font-glyph", diagnostic.Code);
        Assert.Equal("U+10FFFF", diagnostic.CodePoint);
        Assert.Equal("word:runs[4]", diagnostic.Source);
    }

    [Fact]
    public void PdfTextDiagnostics_PlanEmbeddedFontFallbackTextSegmentsCoveredText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));

        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            "Invoice Cafe 123",
            new[] { candidate },
            "word:paragraph[4]");

        PdfTextFallbackSegment segment = Assert.Single(plan.Segments);
        Assert.True(plan.IsFullyCovered);
        Assert.Empty(plan.Diagnostics);
        Assert.Equal("Invoice Cafe 123", plan.OriginalText);
        Assert.Equal("Invoice Cafe 123", segment.Text);
        Assert.Equal(0, segment.StartIndex);
        Assert.Equal("Primary", segment.FontName);
        Assert.Equal(0, segment.FontIndex);
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsMapsFontSlotsAndPreservesStyle() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            "Invoice Cafe 123",
            new[] { candidate },
            "word:paragraph[4]");
        var template = new TextRun(
            "template",
            bold: true,
            underline: true,
            color: PdfColor.FromRgb(1, 2, 3),
            italic: true,
            strike: true,
            fontSize: 14,
            baseline: PdfTextBaseline.Superscript,
            backgroundColor: PdfColor.FromRgb(250, 240, 200));

        TextRun run = Assert.Single(plan.ToTextRuns(new[] { PdfStandardFont.Courier }, template));

        Assert.Equal("Invoice Cafe 123", run.Text);
        Assert.Equal(PdfStandardFont.Courier, run.Font);
        Assert.True(run.Bold);
        Assert.True(run.Underline);
        Assert.True(run.Italic);
        Assert.True(run.Strike);
        Assert.Equal(14, run.FontSize);
        Assert.Equal(PdfTextBaseline.Superscript, run.Baseline);
        Assert.Equal(template.Color, run.Color);
        Assert.Equal(template.BackgroundColor, run.BackgroundColor);
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsKeepsLeadingLayoutControls() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            "\n\tInvoice",
            new[] { candidate },
            "word:paragraph[4]");

        IReadOnlyList<TextRun> runs = plan.ToTextRuns(new[] { PdfStandardFont.Helvetica });

        Assert.Equal(3, runs.Count);
        Assert.Equal("\n", runs[0].Text);
        Assert.Equal("\t", runs[1].Text);
        Assert.Equal("Invoice", runs[2].Text);
        Assert.Equal(PdfStandardFont.Helvetica, runs[2].Font);
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsKeepsLayoutControlsBetweenFallbackSegments() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            "A\nB\tC\rD",
            new[] { candidate },
            "word:paragraph[4]");

        Assert.Collection(
            plan.Segments,
            segment => {
                Assert.Equal("A", segment.Text);
                Assert.Equal(0, segment.StartIndex);
            },
            segment => {
                Assert.Equal("B", segment.Text);
                Assert.Equal(2, segment.StartIndex);
            },
            segment => {
                Assert.Equal("C", segment.Text);
                Assert.Equal(4, segment.StartIndex);
            },
            segment => {
                Assert.Equal("D", segment.Text);
                Assert.Equal(6, segment.StartIndex);
            });

        IReadOnlyList<TextRun> runs = plan.ToTextRuns(new[] { PdfStandardFont.Helvetica });

        Assert.Collection(
            runs,
            run => Assert.Equal("A", run.Text),
            run => Assert.Equal("\n", run.Text),
            run => Assert.Equal("B", run.Text),
            run => Assert.Equal("\t", run.Text),
            run => Assert.Equal("C", run.Text),
            run => Assert.Equal("\n", run.Text),
            run => Assert.Equal("D", run.Text));
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsNormalizesCrLfBetweenFallbackSegments() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            "A\r\nB",
            new[] { candidate },
            "word:paragraph[4]");

        IReadOnlyList<TextRun> runs = plan.ToTextRuns(new[] { PdfStandardFont.Helvetica });

        Assert.Collection(
            runs,
            run => Assert.Equal("A", run.Text),
            run => Assert.Equal("\n", run.Text),
            run => Assert.Equal("B", run.Text));
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsRejectsIncompletePlans() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        string text = "Invoice " + char.ConvertFromUtf32(0x10FFFF);
        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            text,
            new[] { candidate },
            "word:paragraph[4]");

        Assert.Throws<InvalidOperationException>(() => plan.ToTextRuns(new[] { PdfStandardFont.Helvetica }));
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsRequiresFontSlotForEachCandidate() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            "Invoice",
            new[] { candidate },
            "word:paragraph[4]");

        Assert.Throws<ArgumentException>(() => plan.ToTextRuns(new Dictionary<int, PdfStandardFont>()));
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_RegistersStyledSlotsForFallbackFamilies() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesBoldItalic });
        var options = new PdfOptions();

        options.RegisterEmbeddedFontFallbacks(fallbackSet);

        Assert.Contains(PdfStandardFont.TimesRoman, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.TimesBold, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.TimesItalic, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.TimesBoldItalic, options.EmbeddedFonts.Keys);
        Assert.Equal(PdfStandardFont.TimesRoman, fallbackSet.FontSlots[0]);
        Assert.Equal("Primary-Bold", options.EmbeddedFonts[PdfStandardFont.TimesBold].FontName);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_NamedFamiliesCoexistWithAllCompatibilitySlots() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] font = File.ReadAllBytes(fontPath);
        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Portable Fallback", font) });
        var options = new PdfOptions { CompressContentStreams = false };
        options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Sans", font));
        options.RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Document Serif", font));
        options.RegisterFontFamily(PdfStandardFont.Courier, new PdfEmbeddedFontFamily("Document Mono", font));
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Portable Fallback", new byte[] { 0 }));
        options.RegisterEmbeddedFontFallbacks(fallbackSet);

        Assert.True(options.TryResolveNamedFontFace("Portable Fallback", bold: false, italic: false, out PdfNamedFontFace fallbackFace));
        Assert.True(options.TryGetNamedFontProgramForGeneration(fallbackFace, out PdfTrueTypeFontProgram? programBefore));

        IReadOnlyList<TextRun> runs = fallbackSet.PlanTextRuns("Invoice Cafe");
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(runs))
            .ToBytes();
        Assert.True(options.TryGetNamedFontProgramForGeneration(fallbackFace, out PdfTrueTypeFontProgram? programAfterRendering));

        TextRun run = Assert.Single(runs);
        Assert.Same(programBefore, programAfterRendering);
        Assert.True(fallbackSet.UsesNamedFontFamilies);
        Assert.Empty(fallbackSet.FontSlots);
        Assert.Equal("Portable Fallback", run.FontFamily);
        Assert.Null(run.Font);
        Assert.True(options.HasNamedFontFamily("Portable Fallback"));
        Assert.Contains("/BaseFont /PortableFallback-Regular", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
        Assert.Contains("Invoice Cafe", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_PlanTextRunsUsesRegisteredStyledFontSlots() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.Helvetica });
        IReadOnlyList<TextRun> runs = fallbackSet.PlanTextRuns(
            "Invoice Cafe",
            "word:paragraph[8]",
            TextRun.Bolded("template"));

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Paragraph(paragraph => paragraph.Runs(runs))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        TextRun run = Assert.Single(runs);
        Assert.True(run.Bold);
        Assert.Equal(PdfStandardFont.Helvetica, run.Font);
        Assert.Contains("/BaseFont /Primary-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("Invoice Cafe", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_TryPlanTextRunsReturnsRenderableRunsForCoveredText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();

        bool covered = fallbackSet.TryPlanTextRuns(
            "Invoice Cafe",
            out IReadOnlyList<TextRun> runs,
            "word:paragraph[11]",
            TextRun.Italicized("template", fontSize: 12),
            report,
            "OfficeIMO.Word.Pdf");

        TextRun run = Assert.Single(runs);
        Assert.True(covered);
        Assert.True(run.Italic);
        Assert.Equal(12, run.FontSize);
        Assert.Equal(PdfStandardFont.Helvetica, run.Font);
        Assert.False(report.HasWarnings);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_AnalyzeAdvancedTextLayoutReportsSelectedOpenTypeFeaturesWithSourceIndexes() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Source Serif CFF", File.ReadAllBytes(fontPath!)) },
            new[] { PdfStandardFont.TimesRoman });

        IReadOnlyList<PdfTextShapingDiagnostic> diagnostics = fallbackSet.AnalyzeAdvancedTextLayout(
            "\toffice cafe\u0301",
            "word:paragraph[fallback-features]");

        PdfTextShapingDiagnostic ligature = Assert.Single(diagnostics, item => item.Code == "unsupported-font-ligature-substitution");
        PdfTextShapingDiagnostic mark = Assert.Single(diagnostics, item => item.Code == "unsupported-font-mark-positioning");

        Assert.Equal(2, ligature.Index);
        Assert.Equal("U+0066", ligature.CodePoint);
        Assert.Equal("OpenType GSUB ligature", ligature.Script);
        Assert.Equal("word:paragraph[fallback-features]", ligature.Source);
        Assert.Equal("OpenType GPOS mark", mark.Script);
        Assert.Equal("U+0301", mark.CodePoint);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_TryPlanTextRunsReportsOpenTypeFeatureWarnings() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Source Serif CFF", File.ReadAllBytes(fontPath!)) },
            new[] { PdfStandardFont.TimesRoman });
        var report = new PdfConversionReport();

        bool covered = fallbackSet.TryPlanTextRuns(
            "office",
            out IReadOnlyList<TextRun> runs,
            "word:paragraph[fallback-report]",
            report: report,
            converter: "OfficeIMO.Word.Pdf");

        Assert.True(covered);
        Assert.NotEmpty(runs);
        PdfConversionWarning warning = Assert.Single(report.Warnings, item => item.Code == "unsupported-font-ligature-substitution");
        Assert.Equal("OfficeIMO.Word.Pdf", warning.Converter);
        Assert.Equal("word:paragraph[fallback-report]", warning.Source);
        Assert.Equal("OpenType GSUB ligature", warning.Details["script"]);
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_TryPlanTextRunsReportsUncoveredGlyphsWithoutThrowing() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();

        bool covered = fallbackSet.TryPlanTextRuns(
            "Invoice " + char.ConvertFromUtf32(0x10FFFF),
            out IReadOnlyList<TextRun> runs,
            "word:paragraph[12]",
            report: report,
            converter: "OfficeIMO.Word.Pdf");

        PdfConversionWarning warning = Assert.Single(report.Warnings);
        Assert.False(covered);
        Assert.Empty(runs);
        Assert.Equal("missing-embedded-font-fallback-glyph", warning.Code);
        Assert.Equal("OfficeIMO.Word.Pdf", warning.Converter);
        Assert.Equal("word:paragraph[12]", warning.Source);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
    }

    [Fact]
    public void PdfParagraphBuilder_FallbackTextUsesCurrentStyleAndRegisteredFallbackSet() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.Helvetica });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Paragraph(paragraph => paragraph
                .Bold()
                .Italic()
                .Underline()
                .FontSize(13)
                .Color(PdfColor.FromRgb(10, 20, 30))
                .FallbackText(fallbackSet, "Invoice Cafe", "word:paragraph[9]"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/BaseFont /Primary-BoldItalic", raw, StringComparison.Ordinal);
        Assert.Contains("13 Tf", raw, StringComparison.Ordinal);
        Assert.Contains("Invoice Cafe", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksSplitUnsupportedRichTextRunsAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Paragraph(paragraph => paragraph
                .Bold()
                .Text("Zażółć gęślą jaźń"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/BaseFont /Primary-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksSplitUnsupportedHeadingsAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] topLevelBytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .H1("Zażółć gęślą jaźń")
            .ToBytes();
        byte[] rowColumnBytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column.H3("Łódź"))))))
            .ToBytes();

        string topLevelRaw = Encoding.ASCII.GetString(topLevelBytes);
        string rowColumnRaw = Encoding.ASCII.GetString(rowColumnBytes);
        string topLevelText = PdfReadDocument.Open(topLevelBytes).ExtractText();
        string rowColumnText = PdfReadDocument.Open(rowColumnBytes).ExtractText();

        Assert.Contains("/BaseFont /Primary-Bold", topLevelRaw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Primary-Bold", rowColumnRaw, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", topLevelText, StringComparison.Ordinal);
        Assert.Contains("Łódź", rowColumnText, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksSplitUnsupportedHeaderFooterTextAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Header(header => header.AlignCenter().Text("Nagłówek Łódź"))
            .Footer(footer => footer.AlignRight().Text("Stopka Zażółć {page}/{pages}"))
            .Paragraph(paragraph => paragraph.Text("Body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/BaseFont /Primary", raw, StringComparison.Ordinal);
        Assert.Contains("Body", extracted, StringComparison.Ordinal);
        Assert.Contains("łó", extracted, StringComparison.Ordinal);
        Assert.Contains("Stopka Zażółć 1/1", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksSplitUnsupportedTableCaptionsAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });
        var topLevelStyle = TableStyles.Minimal();
        topLevelStyle.Caption = "Tabela Łódź Zażółć";
        topLevelStyle.CaptionFontSize = 10;
        topLevelStyle.CaptionSpacingAfter = 4;
        var rowColumnStyle = TableStyles.Minimal();
        rowColumnStyle.Caption = "Tabela kolumna Łódź";
        rowColumnStyle.CaptionFontSize = 10;
        rowColumnStyle.CaptionSpacingAfter = 4;

        byte[] topLevelBytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Alpha", "1" }
            }, style: topLevelStyle)
            .ToBytes();
        byte[] rowColumnBytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column.Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Beta", "2" }
            }, style: rowColumnStyle))))))
            .ToBytes();

        string topLevelRaw = Encoding.ASCII.GetString(topLevelBytes);
        string rowColumnRaw = Encoding.ASCII.GetString(rowColumnBytes);
        string topLevelText = PdfReadDocument.Open(topLevelBytes).ExtractText();
        string rowColumnText = PdfReadDocument.Open(rowColumnBytes).ExtractText();

        Assert.Contains("/BaseFont /Primary", topLevelRaw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Primary", rowColumnRaw, StringComparison.Ordinal);
        Assert.Contains("Tabela Łódź Zażółć", topLevelText, StringComparison.Ordinal);
        Assert.Contains("Tabela kolumna Łódź", rowColumnText, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksSplitUnsupportedCanvasTextAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Canvas(canvas => canvas
                .Text("Canvas Łódź", 18, 16, 180, 28, fontSize: 12)
                .TextBox("Ramka Zażółć", 18, 54, 190, 42, new PdfCanvasTextBoxStyle {
                    FontSize = 11,
                    PaddingX = 4,
                    PaddingY = 4
                }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/BaseFont /Primary", raw, StringComparison.Ordinal);
        Assert.Contains("Canvas Łódź", extracted, StringComparison.Ordinal);
        Assert.Contains("Ramka Zażółć", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksPreserveExplicitSpacesInCanvasText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string marker = "Café Ω Ж";
        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Canvas(canvas => canvas.Text(marker, 18, 16, 180, 28, fontSize: 12))
            .ToBytes();

        PdfReadDocument read = PdfReadDocument.Open(bytes);

        Assert.Contains(read.Pages[0].GetTextSpans(), span => span.Text == " ");
        Assert.Contains(marker, read.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksRenderFreeTextAnnotationAppearancesAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .FreeTextAnnotation(
                "Komentarz Łódź Zażółć",
                width: 180,
                height: 44,
                fontSize: 11,
                borderColor: PdfColor.Black,
                fillColor: PdfColor.FromRgb(245, 248, 255))
            .Paragraph(paragraph => paragraph.Text("Body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Subtype /FreeText", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Primary", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksRenderFlattenedFreeTextAnnotationAppearancesAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false,
                FlattenVisualAnnotations = true
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .FreeTextAnnotation(
                "Komentarz Łódź Zażółć",
                width: 180,
                height: 44,
                fontSize: 11,
                borderColor: PdfColor.Black,
                fillColor: PdfColor.FromRgb(245, 248, 255))
            .Paragraph(paragraph => paragraph.Text("Body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.DoesNotContain("/Subtype /FreeText", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Primary", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksRenderTextFieldAppearancesAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .TextField("Office.City", value: "Łódź Zażółć", width: 180, height: 24)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);

        Assert.Equal(PdfFormFieldKind.Text, field.Kind);
        Assert.Equal("Łódź Zażółć", field.Value);
        Assert.Equal("Łódź Zażółć", field.DefaultValue);
        Assert.Contains("/FT /Tx", raw, StringComparison.Ordinal);
        Assert.Contains("/V <FEFF", raw, StringComparison.Ordinal);
        Assert.Contains("/DV <FEFF", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Primary", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksRenderChoiceFieldAppearancesAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .ChoiceField(
                "Office.CityChoice",
                new[] { "Łódź", "Zażółć", "Gdańsk" },
                value: "Zażółć",
                width: 180,
                height: 24)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);

        Assert.Equal(PdfFormFieldKind.Choice, field.Kind);
        Assert.Equal("Zażółć", field.Value);
        Assert.Equal("Zażółć", field.DefaultValue);
        Assert.Equal(new[] { "Łódź", "Zażółć", "Gdańsk" }, field.Options.Select(option => option.ExportValue).ToArray());
        Assert.Contains("/FT /Ch", raw, StringComparison.Ordinal);
        Assert.Contains("/Opt [ <FEFF", raw, StringComparison.Ordinal);
        Assert.Contains("/V <FEFF", raw, StringComparison.Ordinal);
        Assert.Contains("/DV <FEFF", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Primary", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfParagraphBuilder_FallbackTextRejectsUncoveredGlyphsBeforeRendering() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.Helvetica });
        string text = "Invoice " + char.ConvertFromUtf32(0x10FFFF);

        Assert.Throws<InvalidOperationException>(() =>
            PdfDocument.Create()
                .RegisterEmbeddedFontFallbacks(fallbackSet)
                .Paragraph(paragraph => paragraph.FallbackText(fallbackSet, text, "word:paragraph[10]")));
    }

    [Fact]
    public void PdfEmbeddedFontFallbackSet_RejectsAmbiguousFontSlotMappings() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidates = new[] {
            new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)),
            new PdfEmbeddedFontFallbackCandidate("Secondary", File.ReadAllBytes(fontPath))
        };

        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFontFallbackSet(
            candidates,
            new[] { PdfStandardFont.Helvetica, PdfStandardFont.HelveticaBold }));
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFontFallbackSet(
            candidates,
            new[] { PdfStandardFont.Helvetica }));
    }

    [Fact]
    public void PdfTextDiagnostics_PlanEmbeddedFontFallbackTextReportsUncoveredGlyphs() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var candidate = new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath));
        string text = "Invoice " + char.ConvertFromUtf32(0x10FFFF) + " tail";

        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
            text,
            new[] { candidate },
            "word:paragraph[5]");

        Assert.False(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);
        Assert.Equal("Invoice ", plan.Segments[0].Text);
        Assert.Equal(" tail", plan.Segments[1].Text);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(plan.Diagnostics);
        Assert.Equal("missing-embedded-font-fallback-glyph", diagnostic.Code);
        Assert.Equal("U+10FFFF", diagnostic.CodePoint);
        Assert.Equal("word:paragraph[5]", diagnostic.Source);
    }

    [Fact]
    public void PdfTextDiagnostics_PlanEmbeddedFontFallbackTextCanSelectLaterCandidate() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        string text = "Emoji 😀 marker";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var candidates = new[] {
                new PdfEmbeddedFontFallbackCandidate("Primary", primary),
                new PdfEmbeddedFontFallbackCandidate("Fallback", File.ReadAllBytes(fallbackPath))
            };
            PdfTextFallbackPlan plan;
            try {
                plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(text, candidates, "word:paragraph[6]");
            } catch (NotSupportedException) {
                continue;
            }

            if (plan.Segments.Any(segment => segment.FontIndex == 1 && segment.Text.Contains("😀", StringComparison.Ordinal))) {
                Assert.True(plan.IsFullyCovered);
                Assert.Empty(plan.Diagnostics);
                return;
            }
        }
    }

    [Fact]
    public void PdfTextFallbackPlan_ToTextRunsCanRenderPlannedFallbackRuns() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        string text = "Emoji 😀 marker";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            byte[] fallback = File.ReadAllBytes(fallbackPath);
            PdfTextFallbackPlan plan;
            try {
                plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText(
                    text,
                    new[] {
                        new PdfEmbeddedFontFallbackCandidate("Primary", primary),
                        new PdfEmbeddedFontFallbackCandidate("Fallback", fallback)
                    },
                    "word:paragraph[7]");
            } catch (NotSupportedException) {
                continue;
            }

            if (!plan.Segments.Any(segment => segment.FontIndex == 1 && segment.Text.Contains("😀", StringComparison.Ordinal))) {
                continue;
            }

            IReadOnlyList<TextRun> runs = plan.ToTextRuns(new[] {
                PdfStandardFont.Helvetica,
                PdfStandardFont.TimesRoman
            });
            Assert.Contains(runs, run => run.Font == PdfStandardFont.TimesRoman && run.Text.Contains("😀", StringComparison.Ordinal));

            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                        CompressContentStreams = false
                    })
                    .EmbedStandardFont(PdfStandardFont.Helvetica, primary, "OfficeIMO Primary Fallback")
                    .EmbedStandardFont(PdfStandardFont.TimesRoman, fallback, "OfficeIMO Emoji Fallback")
                    .Paragraph(paragraph => paragraph.Runs(runs))
                    .ToBytes();
            } catch (ArgumentException exception) when (exception.Message.Contains("not covered by the embedded TrueType font", StringComparison.Ordinal)) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /OfficeIMOPrimaryFallback", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /OfficeIMOEmojiFallback", raw, StringComparison.Ordinal);
            Assert.Contains("Emoji", extracted, StringComparison.Ordinal);
            Assert.Contains("marker", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfTextDiagnostics_AnalyzeGeneratedTextAcceptsSelectedFontPlusFallbackCoverage() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        string text = "Invoice 😀 marker";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", File.ReadAllBytes(fallbackPath)) },
                new[] { PdfStandardFont.TimesRoman });
            var options = new PdfOptions()
                .EmbedStandardFont(PdfStandardFont.Helvetica, primary, "OfficeIMO Primary")
                .RegisterEmbeddedFontFallbacks(fallbackSet);

            try {
                IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(
                    text,
                    options,
                    PdfStandardFont.Helvetica,
                    "word:paragraph[mixed-fallback]");

                if (diagnostics.Count == 0) {
                    return;
                }
            } catch (NotSupportedException) {
                continue;
            }
        }
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksPreserveSelectedFontCoveredSpans() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        string text = "Invoice 😀 marker";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", File.ReadAllBytes(fallbackPath)) },
                new[] { PdfStandardFont.TimesRoman });

            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                        CompressContentStreams = false
                    })
                    .EmbedStandardFont(PdfStandardFont.Helvetica, primary, "OfficeIMO Primary")
                    .RegisterEmbeddedFontFallbacks(fallbackSet)
                    .Paragraph(paragraph => paragraph.Text(text))
                    .ToBytes();
            } catch (Exception exception) when (exception is NotSupportedException || exception is ArgumentException) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /OfficeIMOPrimary", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /EmojiFallback", raw, StringComparison.Ordinal);
            Assert.Contains("Invoice", extracted, StringComparison.Ordinal);
            Assert.Contains("marker", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksSplitMixedUnsupportedTokenWithoutDroppingSelectedText() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        const string text = "A😀B";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", File.ReadAllBytes(fallbackPath)) },
                new[] { PdfStandardFont.TimesRoman });
            if (fallbackSet.PlanText("A").IsFullyCovered ||
                !fallbackSet.PlanText("😀").IsFullyCovered) {
                continue;
            }

            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                        CompressContentStreams = false
                    })
                    .EmbedStandardFont(PdfStandardFont.Helvetica, primary, "OfficeIMO Primary")
                    .RegisterEmbeddedFontFallbacks(fallbackSet)
                    .Paragraph(paragraph => paragraph.Text(text))
                    .ToBytes();
            } catch (Exception exception) when (exception is NotSupportedException || exception is ArgumentException) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /OfficeIMOPrimary", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /EmojiFallback", raw, StringComparison.Ordinal);
            Assert.Contains("A", extracted, StringComparison.Ordinal);
            Assert.Contains("B", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksResolveOnlyUsedCandidates() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        const string text = "Invoice 😀 marker";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            byte[] fallback = File.ReadAllBytes(fallbackPath);
            var coverageProbe = new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", fallback) },
                new[] { PdfStandardFont.TimesRoman });
            if (coverageProbe.PlanText("A").IsFullyCovered ||
                !coverageProbe.PlanText("😀").IsFullyCovered) {
                continue;
            }

            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", fallback),
                    new PdfEmbeddedFontFallbackCandidate("Unused Secondary Fallback", CreateMinimalOpenTypeCffFont()),
                    new PdfEmbeddedFontFallbackCandidate("Unused Tertiary Fallback", CreateMinimalOpenTypeCffFont())
                },
                new[] {
                    PdfStandardFont.Helvetica,
                    PdfStandardFont.TimesRoman,
                    PdfStandardFont.Courier
                });

            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                        CompressContentStreams = false
                    })
                    .EmbedStandardFont(PdfStandardFont.Helvetica, primary, "OfficeIMO Primary")
                    .RegisterEmbeddedFontFallbacks(fallbackSet)
                    .Paragraph(paragraph => paragraph.Text(text))
                    .ToBytes();
            } catch (Exception exception) when (exception is NotSupportedException || exception is ArgumentException) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /OfficeIMOPrimary", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /EmojiFallback", raw, StringComparison.Ordinal);
            Assert.Contains("Invoice", extracted, StringComparison.Ordinal);
            Assert.Contains("marker", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfDocument_RegisteredFallbackReplacementDoesNotOverwriteDocumentDefaultFontSlot() {
        string? primaryPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        const string text = "A😀B";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", File.ReadAllBytes(fallbackPath)) },
                new[] { PdfStandardFont.TimesRoman });
            if (fallbackSet.PlanText("A").IsFullyCovered ||
                !fallbackSet.PlanText("😀").IsFullyCovered) {
                continue;
            }

            byte[] bytes;
            try {
                var options = new PdfOptions {
                    CompressContentStreams = false
                };
                options.RegisterFontFamily(
                    PdfStandardFont.Helvetica,
                    new PdfEmbeddedFontFamily("OfficeIMO Default", primary));
                options.RegisterFontFamily(
                    PdfStandardFont.TimesRoman,
                    new PdfEmbeddedFontFamily("OfficeIMO Primary", primary));
                options.RegisterEmbeddedFontFallbacks(fallbackSet);

                bytes = PdfDocument.Create(options)
                    .Paragraph(paragraph => paragraph.Text("Plain text"))
                    .Paragraph(paragraph => paragraph.Font(PdfStandardFont.TimesRoman).Text(text))
                    .ToBytes();
            } catch (Exception exception) when (exception is NotSupportedException || exception is ArgumentException) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /OfficeIMODefault", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /OfficeIMOPrimary", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /EmojiFallback", raw, StringComparison.Ordinal);
            Assert.Contains("Plain text", extracted, StringComparison.Ordinal);
            Assert.Contains("A", extracted, StringComparison.Ordinal);
            Assert.Contains("B", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksSurviveLaterFontRegistrationInSameSlot() {
        string? primaryPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont() ?? PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryPath == null) {
            return;
        }

        byte[] primary = File.ReadAllBytes(primaryPath);
        const string text = "Invoice \u26A0 marker";
        foreach (string fallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(primaryPath, fallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Issue 2035 Fallback", File.ReadAllBytes(fallbackPath)) },
                new[] { PdfStandardFont.Helvetica });

            byte[] bytes;
            try {
                var options = new PdfOptions {
                    CompressContentStreams = false
                };
                options.RegisterEmbeddedFontFallbacks(fallbackSet);
                options.RegisterFontFamily(
                    PdfStandardFont.Helvetica,
                    new PdfEmbeddedFontFamily("OfficeIMO Primary", primary));

                bytes = PdfDocument.Create(options)
                    .Paragraph(paragraph => paragraph.Bold().Text(text))
                    .ToBytes();
            } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /OfficeIMOPrimary-Bold", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /Issue2035Fallback-Bold", raw, StringComparison.Ordinal);
            Assert.Contains("Invoice", extracted, StringComparison.Ordinal);
            Assert.Contains("marker", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksDoNotReuseSlotsAlreadyEmittedByEarlierRuns() {
        string? textFallbackPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (textFallbackPath == null) {
            return;
        }

        byte[] textFallback = File.ReadAllBytes(textFallbackPath);
        const string polish = "\u0141\u00f3d\u017a";
        foreach (string emojiFallbackPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            if (string.Equals(textFallbackPath, emojiFallbackPath, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            byte[] emojiFallback = File.ReadAllBytes(emojiFallbackPath);
            var fallbackSet = new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate("Issue 2035 Emoji Fallback", emojiFallback),
                    new PdfEmbeddedFontFallbackCandidate("Issue 2035 Polish Fallback", textFallback),
                    new PdfEmbeddedFontFallbackCandidate("Issue 2035 Unused Fallback", CreateMinimalOpenTypeCffFont())
                },
                new[] {
                    PdfStandardFont.Helvetica,
                    PdfStandardFont.TimesRoman,
                    PdfStandardFont.Courier
                });

            PdfTextFallbackPlan polishPlan;
            PdfTextFallbackPlan emojiPlan;
            try {
                polishPlan = fallbackSet.PlanText(polish);
                emojiPlan = fallbackSet.PlanText("\U0001F600");
            } catch (NotSupportedException) {
                continue;
            }

            if (!polishPlan.IsFullyCovered ||
                !polishPlan.Segments.Any(segment => segment.FontIndex == 1) ||
                !emojiPlan.IsFullyCovered ||
                !emojiPlan.Segments.Any(segment => segment.FontIndex == 0)) {
                continue;
            }

            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                        CompressContentStreams = false
                    })
                    .RegisterEmbeddedFontFallbacks(fallbackSet)
                    .Paragraph(paragraph => paragraph.Text("Polish " + polish))
                    .Paragraph(paragraph => paragraph.Text("Emoji \U0001F600"))
                    .ToBytes();
            } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);
            string extracted = PdfReadDocument.Open(bytes).ExtractText();

            Assert.Contains("/BaseFont /Issue2035PolishFallback", raw, StringComparison.Ordinal);
            Assert.Contains("/BaseFont /Issue2035EmojiFallback", raw, StringComparison.Ordinal);
            Assert.Contains("Polish", extracted, StringComparison.Ordinal);
            Assert.Contains("Emoji", extracted, StringComparison.Ordinal);
            return;
        }
    }

    [Fact]
    public void PdfDocument_UseFontFamilyEncodesTextWatermarkWithEmbeddedGlyphs() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Watermark Font", fontPath)
            .Watermark("DRAFT", fontSize: 32, opacity: 0.25)
            .Paragraph(paragraph => paragraph.Text("Watermark font proof"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("<4452414654> Tj", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-Identity-Glyph-UCS", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfRegisteredEmbeddedFontFallbacksSplitUnsupportedTextWatermarksAutomatically() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Primary", File.ReadAllBytes(fontPath)) },
            new[] { PdfStandardFont.TimesRoman });

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .Watermark("DRAFT Łódź", fontSize: 32, opacity: 0.25)
            .Paragraph(paragraph => paragraph.Text("Watermark fallback proof"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        string extracted = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("/BaseFont /Primary-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("DRAFT Łódź", extracted, StringComparison.Ordinal);
        Assert.Contains("Watermark fallback proof", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_UseFontFamilyWrapsLongNonBmpTokensWithoutSplittingSurrogates() {
        foreach (string fontPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                    CompressContentStreams = false,
                    PageWidth = 120,
                    MarginLeft = 24,
                    MarginRight = 24
                })
                .UseFontFamily("OfficeIMO NonBmp Font", fontPath)
                .Paragraph(paragraph => paragraph.Text("AAAAAAAAAAAA😀BBBBBBBBBBBB"))
                .ToBytes();
            } catch (ArgumentException exception) when (exception.Message.Contains("not covered by the embedded TrueType font", StringComparison.Ordinal)) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);

            Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
            Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
            return;
        }
    }

    private static IEnumerable<string> EnumerateLocalNonBmpTrueTypeFonts() {
        string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string[] candidates = {
            Path.Combine(windowsFonts, "seguiemj.ttf"),
            Path.Combine(windowsFonts, "seguisym.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansSymbols2-Regular.ttf"
        };

        foreach (string candidate in candidates) {
            if (File.Exists(candidate)) {
                yield return candidate;
            }
        }
    }

    private static byte[] CreateMultilingualBusinessReport(string fontPath) =>
        PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            })
            .UseFontFamily("OfficeIMO Multilingual", fontPath)
            .Header(header => header.Text("Q2 multilingual revenue report"))
            .Paragraph(paragraph => paragraph.Text("Executive summary: Zażółć gęślą jaźń Łódź."))
            .Paragraph(paragraph => paragraph.Text("Regional notes: Ελλάδα Athens pipeline; Київ renewal forecast."))
            .Footer(footer => footer.Text("Generated proof {page}/{pages}"))
            .ToBytes();

    private static IReadOnlyList<int> ExtractLength1Values(string raw) =>
        Regex.Matches(raw, @"/Length1\s+(\d+)")
            .Cast<Match>()
            .Select(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture))
            .ToArray();

    private static byte[] CreateMinimalOpenTypeCffFont() =>
        new byte[] {
            0x4F, 0x54, 0x54, 0x4F,
            0x00, 0x00,
            0x00, 0x00,
            0x00, 0x00,
            0x00, 0x00
        };

    private static byte[] CreateEmbeddingRestrictedOpenTypeCffFont(byte[] fontData) {
        byte[] restricted = fontData.ToArray();
        int os2Offset = FindOpenTypeTableOffset(restricted, "OS/2");
        WriteUInt16(restricted, os2Offset + 8, 0x0002);
        return restricted;
    }

    private static int CountUnicodeScalars(string text) {
        int count = 0;
        for (int index = 0; index < text.Length; index++) {
            count++;
            if (char.IsHighSurrogate(text[index]) &&
                index + 1 < text.Length &&
                char.IsLowSurrogate(text[index + 1])) {
                index++;
            }
        }

        return count;
    }

    private static int FindObjectIdBefore(string raw, string marker) {
        int markerIndex = raw.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(markerIndex >= 0, "Could not find marker '" + marker + "'.");

        int objectSuffixIndex = raw.LastIndexOf(" obj", markerIndex, StringComparison.Ordinal);
        Assert.True(objectSuffixIndex >= 0, "Could not find object header before marker '" + marker + "'.");

        int lineStart = raw.LastIndexOf('\n', objectSuffixIndex);
        string objectHeader = raw.Substring(lineStart + 1, objectSuffixIndex - lineStart - 1);
        string[] parts = objectHeader.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        Assert.True(parts.Length >= 2, "Could not parse object header '" + objectHeader + "'.");
        return int.Parse(parts[0], CultureInfo.InvariantCulture);
    }

    private static bool TryFindInstalledSystemFontFamily(out PdfEmbeddedFontFamily? family) {
        string[] candidates = {
            "Arial",
            "DejaVu Sans",
            "Liberation Sans",
            "Segoe UI"
        };

        foreach (string candidate in candidates) {
            if (PdfEmbeddedFontFamily.TryFromSystem(candidate, out family) &&
                family != null &&
                (family.Bold != null || family.Italic != null || family.BoldItalic != null)) {
                return true;
            }
        }

        family = null;
        return false;
    }

    private static bool DefaultTextSymbolFallbackFontIsAvailable() {
        string[] candidates = {
            "Segoe UI Symbol",
            "Noto Sans Symbols",
            "Noto Sans Symbols 2",
            "Symbola",
            "DejaVu Sans",
            "Arial Unicode MS",
            "Arial"
        };

        foreach (string candidate in candidates) {
            if (PdfEmbeddedFontFamily.TryFromSystem(candidate, out PdfEmbeddedFontFamily? family) &&
                family != null) {
                return true;
            }
        }

        return false;
    }

    private static bool TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath) {
        string[] candidates = {
            "Arial",
            "DejaVu Sans",
            "Liberation Sans",
            "Segoe UI",
            "Microsoft Sans Serif"
        };

        foreach (string path in EnumerateInstalledTrueTypeFonts()) {
            foreach (string candidate in candidates) {
                if (PdfEmbeddedFontFamily.TryFromSystemFontFiles(candidate, new[] { path }, out PdfEmbeddedFontFamily? family) &&
                    family != null &&
                    family.Bold == null &&
                    family.Italic == null &&
                    family.BoldItalic == null) {
                    familyName = candidate;
                    fontPath = path;
                    return true;
                }
            }
        }

        familyName = string.Empty;
        fontPath = string.Empty;
        return false;
    }

    private static IEnumerable<string> EnumerateInstalledTrueTypeFonts() {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrWhiteSpace(windows)) {
            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(windows, "Fonts"))) {
                yield return font;
            }
        }

        foreach (string font in EnumerateTrueTypeFonts("/usr/share/fonts")) {
            yield return font;
        }

        foreach (string font in EnumerateTrueTypeFonts("/usr/local/share/fonts")) {
            yield return font;
        }

        string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrWhiteSpace(userProfile)) {
            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(userProfile, ".local", "share", "fonts"))) {
                yield return font;
            }

            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(userProfile, ".fonts"))) {
                yield return font;
            }

            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(userProfile, "Library", "Fonts"))) {
                yield return font;
            }
        }

        foreach (string font in EnumerateTrueTypeFonts("/Library/Fonts")) {
            yield return font;
        }

        foreach (string font in EnumerateTrueTypeFonts("/System/Library/Fonts")) {
            yield return font;
        }
    }

    private static IEnumerable<string> EnumerateTrueTypeFonts(string root) {
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) {
            yield break;
        }

        string[] files;
        try {
            files = Directory.GetFiles(root, "*.ttf", SearchOption.AllDirectories);
        } catch (IOException) {
            yield break;
        } catch (UnauthorizedAccessException) {
            yield break;
        }

        foreach (string file in files) {
            yield return file;
        }
    }

    private static byte[] CreateSingleFaceTrueTypeCollection(byte[] fontData) {
        return CreateRepeatedFaceTrueTypeCollection(fontData, 1);
    }

    private static byte[] CreateRepeatedFaceTrueTypeCollection(byte[] fontData, int fontCount) {
        int fontOffset = 12 + fontCount * 4;
        int tableCount = ReadUInt16(fontData, 4);
        int sourceDirectoryLength = 12 + tableCount * 16;
        int collectionLength = fontOffset + fontData.Length;
        byte[] collection = new byte[collectionLength];

        collection[0] = (byte)'t';
        collection[1] = (byte)'t';
        collection[2] = (byte)'c';
        collection[3] = (byte)'f';
        WriteUInt32(collection, 4, 0x00010000);
        WriteUInt32(collection, 8, (uint)fontCount);
        for (int index = 0; index < fontCount; index++) {
            WriteUInt32(collection, 12 + index * 4, (uint)fontOffset);
        }

        Array.Copy(fontData, 0, collection, fontOffset, fontData.Length);

        for (int i = 0; i < tableCount; i++) {
            int recordOffset = fontOffset + 12 + i * 16;
            uint sourceOffset = ReadUInt32(fontData, 12 + i * 16 + 8);
            WriteUInt32(collection, recordOffset + 8, (uint)(fontOffset + sourceOffset));
        }

        Assert.True(sourceDirectoryLength <= fontData.Length);
        return collection;
    }

    private static ushort ReadUInt16(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) |
        ((uint)data[offset + 1] << 16) |
        ((uint)data[offset + 2] << 8) |
        data[offset + 3];

    private static int FindOpenTypeTableOffset(byte[] data, string tag) {
        int tableCount = ReadUInt16(data, 4);
        for (int index = 0; index < tableCount; index++) {
            int recordOffset = 12 + index * 16;
            string candidate = Encoding.ASCII.GetString(data, recordOffset, 4);
            if (string.Equals(candidate, tag, StringComparison.Ordinal)) {
                return checked((int)ReadUInt32(data, recordOffset + 8));
            }
        }

        throw new InvalidOperationException("Required OpenType table '" + tag + "' was not found.");
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
