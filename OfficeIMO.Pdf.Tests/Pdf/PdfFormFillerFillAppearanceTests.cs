using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FillFields_UpdatesSimpleTextAndButtonValues() {
        byte[] filled = PdfFormFiller.FillFields(BuildHierarchicalFormPdf(), new Dictionary<string, string> {
            ["Person.Name"] = "Evotec",
            ["AcceptTerms"] = "Off"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.True(info.HasReadableFormFields);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms" }, info.FormFieldNames);
        Assert.Equal("Evotec", info.FormFields[0].Value);
        Assert.Equal("Off", info.FormFields[1].Value);
        Assert.Contains("/NeedAppearances false", Encoding.ASCII.GetString(filled));
        Assert.Equal(false, info.AcroFormNeedAppearances);
        Assert.False(PdfInspector.Preflight(filled).CanRewrite);
    }

    [Fact]
    public void FillFields_CanKeepNeedAppearancesForLegacyViewers() {
        var options = new PdfFormFillerOptions {
            KeepNeedAppearances = true
        };

        byte[] filled = PdfFormFiller.FillFields(BuildHierarchicalFormPdf(), new Dictionary<string, string> {
            ["Person.Name"] = "Evotec",
            ["AcceptTerms"] = "Off"
        }, options);

        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Contains("/NeedAppearances true", Encoding.ASCII.GetString(filled));
        Assert.Equal(true, info.AcroFormNeedAppearances);
    }

    [Fact]
    public void FillFields_GeneratesSimpleTextWidgetAppearance() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Visible value"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Equal("Visible value", info.FormFields[0].Value);
        Assert.Contains("/Subtype /Form", output);
        Assert.Contains("/AP << /N", output);
        Assert.Contains("/Helv", output);
        Assert.Contains("<56697369626C652076616C7565> Tj", output);
    }

    [Fact]
    public void FlattenFields_SynthesizesRichTextWidgetAppearanceFromRichValue() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildRichTextWidgetFormPdfWithoutAppearance());
        string output = Encoding.ASCII.GetString(flattened);

        Assert.Empty(PdfInspector.Inspect(flattened).FormFields);
        Assert.DoesNotContain("/Subtype /Widget", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.DoesNotContain("<body>", output, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica-Bold", output, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica-Oblique", output, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 12 Tf 0.2 0 0.2 rg", output, StringComparison.Ordinal);
        Assert.Contains("BT /HelvB 12 Tf 0.2 0 0.2 rg", output, StringComparison.Ordinal);
        Assert.Contains("BT /HelvI 10 Tf 0 0.4 0.8 rg", output, StringComparison.Ordinal);
        Assert.Contains("0 0.4 0.8 RG", output, StringComparison.Ordinal);
        Assert.Contains("0.95 0.98 1 rg", output, StringComparison.Ordinal);
        Assert.Contains("[3 2] 0 d", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_GeneratesTextAppearanceWithInheritedQuadding() {
        byte[] filled = PdfFormFiller.FillFields(BuildRightAlignedChildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Person.Name"] = "Right"
        });

        string output = Encoding.ASCII.GetString(filled);
        System.Text.RegularExpressions.Match textPosition = System.Text.RegularExpressions.Regex.Match(
            output,
            @"(?<x>[0-9.]+) [0-9.]+ Td <5269676874> Tj");

        Assert.Contains("<5269676874> Tj", output);
        Assert.True(textPosition.Success);
        Assert.True(double.Parse(textPosition.Groups["x"].Value, System.Globalization.CultureInfo.InvariantCulture) > 3D);
        Assert.DoesNotContain(" 3 10.64 Td <5269676874> Tj", output);
    }

    [Fact]
    public void FillFields_UsesGrayDefaultAppearanceTextColor() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdfWithDefaultAppearance("/Helv 10 Tf 0.25 g"), new Dictionary<string, string> {
            ["Name"] = "Gray"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Gray", field.Value);
        Assert.Contains("0.25 0.25 0.25 rg", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_UsesDefaultAppearanceFontSize() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdfWithDefaultAppearance("/Helv 8 Tf 0.25 g"), new Dictionary<string, string> {
            ["Name"] = "Small"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Small", field.Value);
        Assert.Contains("BT /Helv 8 Tf 0.25 0.25 0.25 rg", output, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 12 Tf 0.25 0.25 0.25 rg", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_PreservesWidgetBorderStyleWidthInGeneratedAppearance() {
        byte[] filled = PdfFormFiller.FillFields(
            BuildTextWidgetFormPdfWithWidgetAppearanceStyle(" /MK << /BC [1 0 0] /BG [0.8 0.9 1] >> /BS << /S /D /W 3 /D [4 2] >>"),
            new Dictionary<string, string> {
                ["Name"] = "Border"
            });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Border", field.Value);
        Assert.Contains("0.8 0.9 1 rg", output, StringComparison.Ordinal);
        Assert.Contains("1 0 0 RG 3 w", output, StringComparison.Ordinal);
        Assert.Contains("[4 2] 0 d", output, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0 0 RG 1 w", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_PreservesWidgetUnderlineBorderStyleInGeneratedAppearance() {
        byte[] filled = PdfFormFiller.FillFields(
            BuildTextWidgetFormPdfWithWidgetAppearanceStyle(" /MK << /BC [1 0 0] /BG [0.8 0.9 1] >> /BS << /S /U /W 2 >>"),
            new Dictionary<string, string> {
                ["Name"] = "Underline"
            });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Underline", field.Value);
        Assert.Contains("1 0 0 RG 2 w 0 1 m 160 1 l S", output, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0 0 RG 2 w 1 1 158 18 re S", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_PreservesWidgetBeveledBorderStyleInGeneratedAppearance() {
        byte[] filled = PdfFormFiller.FillFields(
            BuildTextWidgetFormPdfWithWidgetAppearanceStyle(" /MK << /BC [0 0 0] /BG [0.8 0.9 1] >> /BS << /S /B /W 2 >>"),
            new Dictionary<string, string> {
                ["Name"] = "Beveled"
            });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Beveled", field.Value);
        Assert.Contains("0.55 0.55 0.55 RG 2 w 1 1 m 1 19 l 159 19 l S", output, StringComparison.Ordinal);
        Assert.Contains("0 0 0 RG 2 w 1 1 m 159 1 l 159 19 l S", output, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0 0 RG 2 w 1 1 158 18 re S", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_UsesInheritedDefaultAppearanceFontSizeAndTextColor() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdfWithAcroFormDefaultAppearance("/Helv 8.5 Tf 0.1 0.2 0.3 rg"), new Dictionary<string, string> {
            ["Name"] = "Inherited"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Inherited", field.Value);
        Assert.Contains("BT /Helv 8.5 Tf 0.1 0.2 0.3 rg", output, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 12 Tf 0.1 0.2 0.3 rg", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_UsesInheritedDefaultAppearanceFontResourceName() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdfWithAcroFormDefaultAppearanceFontResource("/F1 8.5 Tf 0.1 0.2 0.3 rg"), new Dictionary<string, string> {
            ["Name"] = "Resource"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Resource", field.Value);
        Assert.Contains("BT /F1 8.5 Tf 0.1 0.2 0.3 rg", output, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 8.5 Tf 0.1 0.2 0.3 rg", output, StringComparison.Ordinal);
        Assert.Contains("/Font << /F1 6 0 R >>", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_WrapsLongMultilineTextWidgetAppearance() {
        byte[] filled = PdfFormFiller.FillFields(BuildMultilineTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Notes"] = "Alpha Bravo Charlie Delta"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.True(field.IsMultiline);
        Assert.Equal("Alpha Bravo Charlie Delta", field.Value);
        Assert.Contains("/AP << /N", output);
        Assert.DoesNotContain("<416C70686120427261766F20436861726C69652044656C7461> Tj", output);
        Assert.True(System.Text.RegularExpressions.Regex.Matches(output, @"BT /Helv 12 Tf .* Tj ET").Count > 1);
    }

    [Fact]
    public void FillFields_PreservesUnicodeTextStringsWhenRewriting() {
        byte[] filled = PdfFormFiller.FillFields(BuildUnicodeFieldNameFormPdf(), new Dictionary<string, string> {
            ["名"] = "Visible value"
        });

        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);
        string output = Encoding.ASCII.GetString(filled);

        Assert.Equal("名", field.Name);
        Assert.Equal("Visible value", field.Value);
        Assert.Contains("/T <FEFF540D>", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_RejectsUnicodeAppearanceValuesWithoutEmbeddedAppearanceFont() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Łódź"
            }));

        Assert.Contains("PDF WinAnsiEncoding", exception.Message, StringComparison.Ordinal);
        Assert.Contains("embedded Unicode fonts are required", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_UsesConfiguredEmbeddedAppearanceFontForUnicodeValueWithoutSourceFont() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFile("OfficeIMO Fill Font", fontPath);

        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Łódź"
        }, options);

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/AP << /N", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", output, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_UsesConfiguredOpenTypeCffAppearanceFontForUnicodeValueWithoutSourceFont() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("OfficeIMO Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Łódź"
        }, options);

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.StartsWith("%PDF-1.6", output, StringComparison.Ordinal);
        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/AP << /N", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", output, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/FontFile2", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "opentype-cff-charstrings-not-subset");
    }

    [Fact]
    public void TextAppearanceBuilder_RejectsUnencodableEmbeddedAppearanceFontSegmentsWithoutWinAnsiFallback() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(
                120,
                20,
                "Łódź\nZażółć",
                10,
                textWidth: 30,
                fontResourceName: "F0",
                encodeTextSegmentHex: _ => null));

        Assert.Contains("cannot be encoded by the selected embedded appearance font", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ReportsConfiguredAppearanceFontOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("OfficeIMO Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "office cafe\u0301"
        }, options);

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("office cafe\u0301", field.Value);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void FillAndFlattenFields_ReportsConfiguredAppearanceFontOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("OfficeIMO Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "office cafe\u0301"
        }, options);

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("office cafe", extracted, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void FillFields_ThrowsWhenConfiguredAppearanceFontCannotBeUsed() {
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("Broken Fill Font", new byte[] { 0, 1, 2, 3 });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Łódź"
            }, options));

        Assert.Contains("configured appearance font", exception.Message, StringComparison.Ordinal);
        Assert.Contains("Broken Fill Font", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ReportsConfiguredAppearanceFontMissingGlyphDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFile("OfficeIMO Fill Font", fontPath)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Invoice " + char.ConvertFromUtf32(0x10FFFF)
            }, options));

        PdfConversionWarning warning = Assert.Single(report.Warnings);
        Assert.Contains("configured appearance font", exception.Message, StringComparison.Ordinal);
        Assert.Equal("OfficeIMO.Tests", warning.Converter);
        Assert.Equal("missing-embedded-font-glyph", warning.Code);
        Assert.Equal("form field 'Name' appearance", warning.Source);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("U+10FFFF", warning.Details["codePoint"]);
    }

    [Fact]
    public void FillFields_UsesAppearanceFontFallbackSetWhenConfiguredFontCannotBeUsed() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill Font", System.IO.File.ReadAllBytes(fontPath))
            },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("Broken Fill Font", new byte[] { 0, 1, 2, 3 })
            .UseAppearanceFontFallbacks(fallbackSet)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Łódź"
        }, options);

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/Helv0", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_UsesOpenTypeCffAppearanceFontFallbackSetWhenConfiguredFontCannotBeUsed() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("Broken Fill Font", new byte[] { 0, 1, 2, 3 })
            .UseAppearanceFontFallbacks(fallbackSet)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Łódź"
        }, options);

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.StartsWith("%PDF-1.6", output, StringComparison.Ordinal);
        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/Helv0", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/FontFile2", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "opentype-cff-charstrings-not-subset");
    }

    [Fact]
    public void FillFields_ReportsAppearanceFontFallbackOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFallbacks(fallbackSet)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "office cafe\u0301"
        }, options);

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("office cafe\u0301", field.Value);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void FillAndFlattenFields_ReportsAppearanceFontFallbackOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFallbacks(fallbackSet)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "office cafe\u0301"
        }, options);

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("office cafe", extracted, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void FillAndFlattenFields_UsesAppearanceFontFallbackSetForExtractableUnicodeValue() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill Font", System.IO.File.ReadAllBytes(fontPath))
            },
            new[] { PdfStandardFont.Helvetica });
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFallbacks(fallbackSet);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Łódź"
        }, options);

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/Helv0", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", output, StringComparison.Ordinal);
        Assert.Contains("Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ThrowsWhenAppearanceFontFallbackSetCannotCoverValue() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill Font", System.IO.File.ReadAllBytes(fontPath))
            },
            new[] { PdfStandardFont.Helvetica });
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFallbacks(fallbackSet);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Invoice " + char.ConvertFromUtf32(0x10FFFF)
            }, options));

        Assert.Contains("missing-embedded-font-fallback-glyph", exception.Message, StringComparison.Ordinal);
        Assert.Contains("U+10FFFF", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ReportsAppearanceFontFallbackMissingGlyphDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill Font", System.IO.File.ReadAllBytes(fontPath))
            },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFallbacks(fallbackSet)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Invoice " + char.ConvertFromUtf32(0x10FFFF)
            }, options));

        PdfConversionWarning warning = Assert.Single(report.Warnings);
        Assert.Contains("missing-embedded-font-fallback-glyph", exception.Message, StringComparison.Ordinal);
        Assert.Equal("OfficeIMO.Tests", warning.Converter);
        Assert.Equal("missing-embedded-font-fallback-glyph", warning.Code);
        Assert.Equal("form field 'Name' appearance", warning.Source);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("U+10FFFF", warning.Details["codePoint"]);
    }

    [Fact]
    public void FillAndFlattenFields_UsesConfiguredEmbeddedAppearanceFontForExtractableUnicodeValueWithoutSourceFont() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFile("OfficeIMO Fill Font", fontPath);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Łódź"
        }, options);

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.Contains("Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FluentFill_UsesConfiguredEmbeddedAppearanceFontForUnicodeValueWithoutSourceFont() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFile("OfficeIMO Fill Font", fontPath);

        byte[] filled = PdfDocument
            .Open(BuildTextWidgetFormPdf())
            .Forms
            .Fill(new Dictionary<string, string> {
                ["Name"] = "Łódź"
            }, options)
            .ToBytes();

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ReusesInheritedEmbeddedTextAppearanceFontForCoveredUnicodeValue() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdf("Łódź Zażółć");
        if (source.Length == 0) {
            return;
        }

        byte[] filled = PdfFormFiller.FillFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/AP << /N", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_DiscoversNonHelvInheritedEmbeddedTextAppearanceFontForCoveredUnicodeValue() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdfWithResourceName("Łódź Zażółć", "Funi");
        if (source.Length == 0) {
            return;
        }

        byte[] filled = PdfFormFiller.FillFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/AP << /N", output, StringComparison.Ordinal);
        Assert.Contains("/Funi", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ReusesExistingAppearanceEmbeddedFontWhenAcroFormDefaultResourceFontsAreUnavailable() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdfWithAppearanceOnlyFontResource("Łódź Zażółć");
        if (source.Length == 0) {
            return;
        }

        byte[] filled = PdfFormFiller.FillFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/AP << /N", output, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillAndFlattenFields_ReusesInheritedEmbeddedTextAppearanceFontForCoveredUnicodeValue() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdf("Łódź Zażółć");
        if (source.Length == 0) {
            return;
        }

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.Contains("Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FillAndFlattenFields_ReusesExistingAppearanceEmbeddedFontWhenAcroFormDefaultResourceFontsAreUnavailable() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdfWithAppearanceOnlyFontResource("Łódź Zażółć");
        if (source.Length == 0) {
            return;
        }

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.Contains("Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FillAndFlattenFields_DiscoversNonHelvInheritedEmbeddedTextAppearanceFontForCoveredUnicodeValue() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdfWithResourceName("Łódź Zażółć", "Funi");
        if (source.Length == 0) {
            return;
        }

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/Funi", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.Contains("Łódź", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FillAndFlattenFields_ReusesInheritedEmbeddedFontForMultilineUnicodeTextField() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdf(
            "Łódź\nZażółć",
            new PdfFormFieldStyle {
                IsMultiline = true
            });
        if (source.Length == 0) {
            return;
        }

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź\nZażółć"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
        Assert.Contains("Łódź", extracted, StringComparison.Ordinal);
        Assert.Contains("Zażółć", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ReusesInheritedEmbeddedFontForCombUnicodeTextField() {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdf(
            "Łódź",
            new PdfFormFieldStyle {
                IsComb = true,
                MaxLength = 4
            });
        if (source.Length == 0) {
            return;
        }

        byte[] filled = PdfFormFiller.FillFields(source, new Dictionary<string, string> {
            ["Office.City"] = "Łódź"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.True(field.IsComb);
        Assert.Equal(4, field.MaxLength);
        Assert.Equal("Łódź", field.Value);
        Assert.Contains("/Encoding /Identity-H", output, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Type1 /BaseFont /Helvetica", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_LimitsInheritedCidWidthRangeExpansion() {
        byte[] filled = PdfFormFiller.FillFields(BuildType0TextWidgetFormPdfWithWideCidWidths(), new Dictionary<string, string> {
            ["Name"] = "B"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("B", field.Value);
        Assert.Contains("/Subtype /Type0", output, StringComparison.Ordinal);
        Assert.Contains("<1388> Tj", output, StringComparison.Ordinal);
        Assert.Contains("185 12.64 Td <1388> Tj", output, StringComparison.Ordinal);
        Assert.DoesNotContain("194 12.64 Td <1388> Tj", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_GeneratesSimpleButtonWidgetAppearances() {
        byte[] filled = PdfFormFiller.FillFields(BuildCheckboxWidgetWithoutAppearancePdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Equal("Yes", info.FormFields[0].Value);
        Assert.Contains("/AS /Yes", output);
        Assert.Contains("/AP << /N <<", output);
        Assert.Contains("/Off", output);
        Assert.Contains("/Yes", output);
        Assert.Contains("1.25 w", output);
    }

    [Fact]
    public void FillFields_GeneratesRadioButtonWidgetAppearances() {
        byte[] filled = PdfFormFiller.FillFields(BuildRadioWidgetGroupWithoutOffAppearancePdf(), new Dictionary<string, string> {
            ["Payment.Method"] = "Wire"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Wire", field.Value);
        Assert.Contains(field.Widgets, widget => widget.AppearanceState == "Wire");
        Assert.Equal(2, field.Widgets.Count(widget => widget.AppearanceState == "Off"));
        Assert.Contains(" c S", output, StringComparison.Ordinal);
        Assert.DoesNotContain("1.25 w", output, StringComparison.Ordinal);
    }

    private static byte[] BuildEmbeddedUnicodeTextFieldPdf(string value, PdfFormFieldStyle? style = null) {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return Array.Empty<byte>();
        }

        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Fill Font", fontPath)
            .TextField("Office.City", value: value, width: 180, height: style?.IsMultiline == true ? 48 : 24, style: style)
            .ToBytes();
    }

    private static byte[] BuildType0TextWidgetFormPdfWithWideCidWidths() {
        const string toUnicode = """
beginbfchar
<1388> <0042>
endbfchar
""";
        int toUnicodeLength = Encoding.ASCII.GetByteCount(toUnicode);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] /Q 2 /DR << /Font << /Helv 9 0 R >> >> >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /V (A) /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 220 120] /F 4 >>",
            "endobj",
            "9 0 obj",
            "<< /Type /Font /Subtype /Type0 /BaseFont /OfficeIMOTest /Encoding /Identity-H /DescendantFonts [11 0 R] /ToUnicode 10 0 R >>",
            "endobj",
            "10 0 obj",
            $"<< /Length {toUnicodeLength.ToString(CultureInfo.InvariantCulture)} >>",
            "stream",
            toUnicode,
            "endstream",
            "endobj",
            "11 0 obj",
            "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /OfficeIMOTest /DW 1000 /W [0 100000 250] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 12 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static void AssertOpenTypeFeatureAppearanceDiagnostics(PdfConversionReport report) {
        PdfConversionWarning ligature = Assert.Single(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
        PdfConversionWarning mark = Assert.Single(report.Warnings, warning => warning.Code == "unsupported-font-mark-positioning");
        Assert.Equal("OfficeIMO.Tests", ligature.Converter);
        Assert.Equal("form field 'Name' appearance", ligature.Source);
        Assert.Equal("OpenType GSUB ligature", ligature.Details["script"]);
        Assert.Equal("U+0066", ligature.Details["codePoint"]);
        Assert.Equal("OpenType GPOS mark", mark.Details["script"]);
        Assert.Equal("U+0301", mark.Details["codePoint"]);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "opentype-cff-charstrings-not-subset");
    }

    private static byte[] BuildEmbeddedUnicodeTextFieldPdfWithResourceName(string value, string resourceName) {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdf(value);
        if (source.Length == 0) {
            return source;
        }

        return ReplaceAsciiTokenSameLength(source, "/Helv", "/" + resourceName);
    }

    private static byte[] BuildEmbeddedUnicodeTextFieldPdfWithAppearanceOnlyFontResource(string value) {
        byte[] source = BuildEmbeddedUnicodeTextFieldPdf(value);
        if (source.Length == 0) {
            return source;
        }

        return ReplaceFirstAsciiTokenSameLength(source, "/DR << /Font << /Helv ", "/DR << /F0nt << /Helv ");
    }

    private static byte[] ReplaceFirstAsciiTokenSameLength(byte[] source, string oldValue, string newValue) {
        byte[] oldBytes = Encoding.ASCII.GetBytes(oldValue);
        byte[] newBytes = Encoding.ASCII.GetBytes(newValue);
        Assert.Equal(oldBytes.Length, newBytes.Length);

        byte[] result = (byte[])source.Clone();
        for (int index = 0; index <= result.Length - oldBytes.Length; index++) {
            bool matched = true;
            for (int offset = 0; offset < oldBytes.Length; offset++) {
                if (result[index + offset] != oldBytes[offset]) {
                    matched = false;
                    break;
                }
            }

            if (!matched) {
                continue;
            }

            for (int offset = 0; offset < newBytes.Length; offset++) {
                result[index + offset] = newBytes[offset];
            }

            return result;
        }

        Assert.Fail("The PDF fixture did not contain the expected ASCII token.");
        return result;
    }

    private static byte[] ReplaceAsciiTokenSameLength(byte[] source, string oldValue, string newValue) {
        byte[] oldBytes = Encoding.ASCII.GetBytes(oldValue);
        byte[] newBytes = Encoding.ASCII.GetBytes(newValue);
        Assert.Equal(oldBytes.Length, newBytes.Length);

        byte[] result = (byte[])source.Clone();
        for (int index = 0; index <= result.Length - oldBytes.Length; index++) {
            bool matched = true;
            for (int offset = 0; offset < oldBytes.Length; offset++) {
                if (result[index + offset] != oldBytes[offset]) {
                    matched = false;
                    break;
                }
            }

            if (!matched) {
                continue;
            }

            for (int offset = 0; offset < newBytes.Length; offset++) {
                result[index + offset] = newBytes[offset];
            }
        }

        return result;
    }
}
