using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfCore.PdfOptions CreatePdfOptions(PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options) {
        PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
        pdfOptions.ReportDiagnosticsTo(options.ConversionReport, "OfficeIMO.PowerPoint.Pdf");

        pdfOptions.PageWidth = presentation.SlideSize.WidthPoints;
        pdfOptions.PageHeight = presentation.SlideSize.HeightPoints;
        pdfOptions.Margins = PdfCore.PageMargins.Uniform(0);
        bool preserveConfiguredFontSlots = options.PdfOptions != null;
        if (!string.IsNullOrWhiteSpace(options.FontFamily) &&
            TryApplyPdfFontFamily(options.FontFamily, pdfOptions)) {
            preserveConfiguredFontSlots = true;
        }

        RegisterPresentationFonts(pdfOptions, presentation, options, preserveConfiguredFontSlots);
        ApplyDefaultEmbeddedFontFallback(pdfOptions, options, preserveConfiguredFontSlots);
        return pdfOptions;
    }

    private static void ApplyDefaultEmbeddedFontFallback(PdfCore.PdfOptions pdfOptions, PowerPointPdfSaveOptions options, bool preserveConfiguredFontSlots) {
        if (options.PdfOptions == null &&
            !preserveConfiguredFontSlots &&
            !pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
            pdfOptions.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true);
        }
    }

    private static bool TryApplyPdfFontFamily(string? familyName, PdfCore.PdfOptions pdfOptions, bool requireEmbeddedFont = false) {
        return pdfOptions.TryUseOfficeFontFamily(familyName, embedSystemFont: true, requireEmbeddedFont: requireEmbeddedFont);
    }

    private static void RegisterPresentationFonts(PdfCore.PdfOptions pdfOptions, PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options, bool preserveConfiguredFontSlots) {
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
            RegisterPresentationShapesFonts(slide.GetInheritedShapesForExport(), slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth: 0);
            RegisterPresentationShapesFonts(slide.Shapes, slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth: 0);
        }
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
                RegisterPresentationTextBoxFonts(textBox, pdfOptions, registeredFamilies, registeredFontSlots);
            }
            return;
        }

        if (shape is PptCore.PowerPointTable table) {
            if (options.IncludeTables) {
                RegisterPresentationTableFonts(table, pdfOptions, registeredFamilies, registeredFontSlots);
            }
            return;
        }

        if (shape is PptCore.PowerPointGroupShape groupShape && groupShape.OwnerSlide != null) {
            if (options.MaxGroupShapeDepth < 0 || groupDepth < options.MaxGroupShapeDepth) {
                RegisterPresentationShapesFonts(groupShape.OwnerSlide.GetGroupChildren(groupShape), slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options, groupDepth + 1);
            }
        }
    }

    private static void RegisterPresentationTextBoxFonts(PptCore.PowerPointTextBox textBox, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
        RegisterPresentationFontCandidate(textBox.FontName, pdfOptions, registeredFamilies, registeredFontSlots);
        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            foreach (PptCore.PowerPointTextRun run in paragraph.Runs) {
                RegisterPresentationFontCandidate(run.FontName, pdfOptions, registeredFamilies, registeredFontSlots);
            }
        }
    }

    private static void RegisterPresentationTableFonts(PptCore.PowerPointTable table, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
        for (int row = 0; row < table.Rows; row++) {
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell cell = table.GetCell(row, column);
                if (!cell.IsMergedCell) {
                    RegisterPresentationFontCandidate(cell.FontName, pdfOptions, registeredFamilies, registeredFontSlots);
                    RegisterPresentationTableCellRunFonts(cell, pdfOptions, registeredFamilies, registeredFontSlots);
                }
            }
        }
    }

    private static void RegisterPresentationTableCellRunFonts(PptCore.PowerPointTableCell cell, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
        A.TextBody? textBody = cell.Cell.TextBody;
        if (textBody == null) {
            return;
        }

        foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
            foreach (A.Run run in paragraph.Elements<A.Run>()) {
                RegisterPresentationFontCandidate(ReadRunFontName(run.RunProperties), pdfOptions, registeredFamilies, registeredFontSlots);
            }
        }
    }

    private static void RegisterPresentationFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
        if (PdfCore.PdfOptions.TryAddOfficeFontFamilyKey(familyName, registeredFamilies, normalizeKey: null, out string trimmedFamilyName)) {
            pdfOptions.TryRegisterMappedOfficeFontFamily(trimmedFamilyName, registeredFontSlots, embedSystemFont: true, out _);
        }
    }
}
