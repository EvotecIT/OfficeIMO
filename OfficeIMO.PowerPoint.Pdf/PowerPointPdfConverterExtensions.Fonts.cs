using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfCore.PdfOptions CreatePdfOptions(PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options) {
        PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
        pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.PowerPoint.Pdf");

        if (options.PageLayout == PowerPointPdfPageLayout.NotesPages) {
            pdfOptions.PageWidth = 612D;
            pdfOptions.PageHeight = 792D;
        } else if (options.PageLayout == PowerPointPdfPageLayout.Handouts) {
            pdfOptions.PageWidth = 792D;
            pdfOptions.PageHeight = 612D;
        } else {
            pdfOptions.PageWidth = presentation.SlideSize.WidthPoints;
            pdfOptions.PageHeight = presentation.SlideSize.HeightPoints;
        }
        pdfOptions.Margins = PdfCore.PageMargins.Uniform(0);
        bool preserveConfiguredFontSlots = options.PdfOptions != null;
        if (!string.IsNullOrWhiteSpace(options.FontFamily) &&
            TryApplyPdfFontFamily(options.FontFamily, pdfOptions, options.ResourcePolicy.AllowSystemFontEmbedding)) {
            preserveConfiguredFontSlots = true;
        }

        HashSet<PdfCore.PdfStandardFont> registeredFontSlots = RegisterPresentationFonts(pdfOptions, presentation, options, preserveConfiguredFontSlots);
        ApplyTextFallbacks(pdfOptions, options, preserveConfiguredFontSlots, registeredFontSlots);
        return pdfOptions;
    }

    private static void ApplyTextFallbacks(
        PdfCore.PdfOptions pdfOptions,
        PowerPointPdfSaveOptions options,
        bool preserveConfiguredFontSlots,
        IEnumerable<PdfCore.PdfStandardFont> reservedFontSlots) {
        if (!options.ResourcePolicy.AllowSystemFontEmbedding ||
            options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) {
            return;
        }

        PdfCore.PdfTextFallbackFeatures fallbackFeatures = options.TextFallbacks;
        if (preserveConfiguredFontSlots || pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
            fallbackFeatures &= ~PdfCore.PdfTextFallbackFeatures.DocumentFont;
        }

        pdfOptions.UseTextFallbacks(fallbackFeatures, reservedFontSlots, options.ResourcePolicy.AllowSystemFontEmbedding);
    }

    private static bool TryApplyPdfFontFamily(string? familyName, PdfCore.PdfOptions pdfOptions, bool embedSystemFont, bool requireEmbeddedFont = false) {
        return pdfOptions.TryUseOfficeFontFamily(familyName, embedSystemFont, requireEmbeddedFont);
    }

    private static HashSet<PdfCore.PdfStandardFont> RegisterPresentationFonts(PdfCore.PdfOptions pdfOptions, PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options, bool preserveConfiguredFontSlots) {
        var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(preserveConfiguredFontSlots);
        double pageWidth = presentation.SlideSize.WidthPoints;
        double pageHeight = presentation.SlideSize.HeightPoints;
        IReadOnlyList<PptCore.PowerPointSlide> slides = presentation.Slides;
        for (int slideIndex = 0; slideIndex < slides.Count; slideIndex++) {
            PptCore.PowerPointSlide slide = slides[slideIndex];
            if (!options.IncludeHiddenSlides && slide.Hidden) {
                continue;
            }

            int slideNumber = slideIndex + 1;
            IReadOnlyList<PptCore.PowerPointShape> inheritedShapes = slide.GetInheritedShapesForExport();
            bool usesThemeFont =
                ContainsPresentationThemeFontUsage(inheritedShapes, options, groupDepth: 0) ||
                ContainsPresentationThemeFontUsage(slide.Shapes, options, groupDepth: 0);

            RegisterPresentationShapesFonts(inheritedShapes, slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth: 0);
            RegisterPresentationShapesFonts(slide.Shapes, slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth: 0);
            if (usesThemeFont) {
                RegisterPresentationFontCandidate(
                    PptCore.PowerPointTextDefaults.ResolveBodyLatinFont(slide),
                    pdfOptions,
                    registeredFamilies,
                    registeredFontSlots,
                    options,
                    slideNumber,
                    reportSubstitution: !preserveConfiguredFontSlots);
            }
        }

        return registeredFontSlots;
    }

    private static bool ContainsPresentationThemeFontUsage(
        IReadOnlyList<PptCore.PowerPointShape> shapes,
        PowerPointPdfSaveOptions options,
        int groupDepth) {
        if (!string.IsNullOrWhiteSpace(options.FontFamily)) {
            return false;
        }

        foreach (PptCore.PowerPointShape shape in shapes) {
            if (shape.Hidden) {
                continue;
            }
            if (options.IncludeTextBoxes &&
                shape is PptCore.PowerPointTextBox textBox &&
                TextBoxUsesPresentationThemeFont(textBox)) {
                    return true;
            }
            if (options.IncludeTables &&
                shape is PptCore.PowerPointTable table &&
                TableUsesPresentationThemeFont(table)) {
                    return true;
            }
            if (shape is PptCore.PowerPointGroupShape groupShape &&
                groupShape.OwnerSlide != null &&
                (options.MaxGroupShapeDepth < 0 || groupDepth < options.MaxGroupShapeDepth) &&
                ContainsPresentationThemeFontUsage(
                    groupShape.OwnerSlide.GetGroupChildren(groupShape),
                    options,
                    groupDepth + 1)) {
                return true;
            }
        }

        return false;
    }

    private static bool TextBoxUsesPresentationThemeFont(PptCore.PowerPointTextBox textBox) {
        if (string.IsNullOrWhiteSpace(textBox.Text) ||
            !string.IsNullOrWhiteSpace(textBox.FontName)) {
            return false;
        }

        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            if (!string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered) {
                return true;
            }

            IReadOnlyList<PptCore.PowerPointTextRun> runs = paragraph.Runs;
            if (runs.Count == 0 && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                return true;
            }

            foreach (PptCore.PowerPointTextRun run in runs) {
                if (!string.IsNullOrEmpty(run.Text) && string.IsNullOrWhiteSpace(run.FontName)) {
                    return true;
                }
            }

            foreach (A.Field field in paragraph.Paragraph.Elements<A.Field>()) {
                if (!string.IsNullOrWhiteSpace(field.Text?.Text ?? field.InnerText)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool TableUsesPresentationThemeFont(PptCore.PowerPointTable table) {
        for (int row = 0; row < table.Rows; row++) {
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell cell = table.GetCell(row, column);
                if (cell.IsMergedCell ||
                    string.IsNullOrWhiteSpace(cell.Text) ||
                    !string.IsNullOrWhiteSpace(cell.FontName)) {
                    continue;
                }

                A.TextBody? textBody = cell.Cell.TextBody;
                if (textBody == null) {
                    return true;
                }

                bool hasTextRun = false;
                foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
                    foreach (A.Run run in paragraph.Elements<A.Run>()) {
                        if (!string.IsNullOrWhiteSpace(run.InnerText)) {
                            hasTextRun = true;
                            if (string.IsNullOrWhiteSpace(ReadRunFontName(run.RunProperties))) {
                                return true;
                            }
                        }
                    }

                    foreach (A.Field field in paragraph.Elements<A.Field>()) {
                        if (!string.IsNullOrWhiteSpace(field.Text?.Text ?? field.InnerText)) {
                            return true;
                        }
                    }
                }

                if (!hasTextRun) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void RegisterPresentationShapesFonts(IReadOnlyList<PptCore.PowerPointShape> shapes, int slideNumber, double pageWidth, double pageHeight, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options, int groupDepth) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            RegisterPresentationShapeFonts(shape, slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth);
        }
    }

    private static void RegisterPresentationShapeFonts(PptCore.PowerPointShape shape, int slideNumber, double pageWidth, double pageHeight, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options, int groupDepth) {
        if (shape.Hidden) {
            return;
        }

        if (!TryGetShapeBox(shape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: false, out _, out _, out _, out _)) {
            return;
        }

        if (shape is PptCore.PowerPointTextBox textBox) {
            if (options.IncludeTextBoxes) {
                RegisterPresentationTextBoxFonts(textBox, slideNumber, pdfOptions, registeredFamilies, registeredFontSlots, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointTable table) {
            if (options.IncludeTables) {
                RegisterPresentationTableFonts(table, slideNumber, pdfOptions, registeredFamilies, registeredFontSlots, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointGroupShape groupShape && groupShape.OwnerSlide != null) {
            if (options.MaxGroupShapeDepth < 0 || groupDepth < options.MaxGroupShapeDepth) {
                RegisterPresentationShapesFonts(groupShape.OwnerSlide.GetGroupChildren(groupShape), slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth + 1);
            }
        }
    }

    private static void RegisterPresentationTextBoxFonts(PptCore.PowerPointTextBox textBox, int slideNumber, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options) {
        RegisterPresentationFontCandidate(textBox.FontName, pdfOptions, registeredFamilies, registeredFontSlots, options, slideNumber);
        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            foreach (PptCore.PowerPointTextRun run in paragraph.Runs) {
                RegisterPresentationFontCandidate(run.FontName, pdfOptions, registeredFamilies, registeredFontSlots, options, slideNumber);
            }
        }
    }

    private static void RegisterPresentationTableFonts(PptCore.PowerPointTable table, int slideNumber, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options) {
        for (int row = 0; row < table.Rows; row++) {
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell cell = table.GetCell(row, column);
                if (!cell.IsMergedCell) {
                    RegisterPresentationFontCandidate(cell.FontName, pdfOptions, registeredFamilies, registeredFontSlots, options, slideNumber);
                    RegisterPresentationTableCellRunFonts(cell, slideNumber, pdfOptions, registeredFamilies, registeredFontSlots, options);
                }
            }
        }
    }

    private static void RegisterPresentationTableCellRunFonts(PptCore.PowerPointTableCell cell, int slideNumber, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options) {
        A.TextBody? textBody = cell.Cell.TextBody;
        if (textBody == null) {
            return;
        }

        foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
            foreach (A.Run run in paragraph.Elements<A.Run>()) {
                RegisterPresentationFontCandidate(ReadRunFontName(run.RunProperties), pdfOptions, registeredFamilies, registeredFontSlots, options, slideNumber);
            }
        }
    }

    private static void RegisterPresentationFontCandidate(
        string? familyName,
        PdfCore.PdfOptions pdfOptions,
        HashSet<string> registeredFamilies,
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
        PowerPointPdfSaveOptions options,
        int slideNumber,
        bool reportSubstitution = true) {
        if (PdfCore.PdfOptions.TryAddOfficeFontFamilyKey(familyName, registeredFamilies, normalizeKey: null, out string trimmedFamilyName)) {
            if (pdfOptions.HasNamedFontFamily(trimmedFamilyName)) {
                return;
            }

            bool embedSystemFont = options.ResourcePolicy.AllowSystemFontEmbedding;
            if (embedSystemFont && pdfOptions.TryRegisterNamedOfficeFontFamily(trimmedFamilyName, out _)) {
                return;
            }

            bool mapped = pdfOptions.TryRegisterMappedOfficeFontFamily(
                trimmedFamilyName,
                registeredFontSlots,
                embedSystemFont,
                out PdfCore.PdfStandardFont fallback);
            bool representedExactly =
                mapped &&
                (pdfOptions.EmbeddedFontFamilySlotMatches(fallback, trimmedFamilyName) ||
                 (!pdfOptions.HasEmbeddedStandardFontFamily(fallback) &&
                  PdfCore.PdfStandardFontMapper.IsStandardPdfFamilyEquivalent(trimmedFamilyName, fallback)));
            if (reportSubstitution && !representedExactly) {
                PdfCore.PdfStandardFont reportedFallback = mapped
                    ? fallback
                    : PdfCore.PdfStandardFont.Helvetica;
                PdfCore.PdfStandardFont normalizedFallback = PdfCore.PdfStandardFontMapper.GetFontFamily(reportedFallback);
                options.Report.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.PowerPoint.Pdf",
                    "font-family-substitution",
                    "Slide " + slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    "The source font family '" + trimmedFamilyName + "' was unavailable or could not be embedded; generated text uses the mapped PDF family " + normalizedFallback + ".",
                    details: new Dictionary<string, string> {
                        ["slideNumber"] = slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ["fontFamily"] = trimmedFamilyName,
                        ["fallbackSlot"] = normalizedFallback.ToString()
                    }));
            }
        }
    }
}
