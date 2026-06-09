using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
            !HasEmbeddedFontSlot(pdfOptions, pdfOptions.DefaultFont)) {
            TryApplyPdfFontFamily(DefaultEmbeddedFontFamily, pdfOptions, requireEmbeddedFont: true);
        }
    }

    private static bool TryApplyPdfFontFamily(string? familyName, PdfCore.PdfOptions pdfOptions, bool requireEmbeddedFont = false) {
        if (string.IsNullOrWhiteSpace(familyName)) {
            return false;
        }

        PdfCore.PdfStandardFont beforeDefault = pdfOptions.DefaultFont;
        PdfCore.PdfStandardFont beforeHeader = pdfOptions.HeaderFont;
        PdfCore.PdfStandardFont beforeFooter = pdfOptions.FooterFont;
        string beforeEmbeddedFonts = CaptureEmbeddedFontState(pdfOptions);
        pdfOptions.UseOfficeFontFamily(familyName);

        bool changed = beforeDefault != pdfOptions.DefaultFont ||
                       beforeHeader != pdfOptions.HeaderFont ||
                       beforeFooter != pdfOptions.FooterFont ||
                       !string.Equals(beforeEmbeddedFonts, CaptureEmbeddedFontState(pdfOptions), StringComparison.Ordinal);
        return changed && (!requireEmbeddedFont || HasEmbeddedFontSlot(pdfOptions, pdfOptions.DefaultFont));
    }

    private static string CaptureEmbeddedFontState(PdfCore.PdfOptions pdfOptions) {
        return string.Join("|", pdfOptions.EmbeddedFonts
            .OrderBy(font => font.Key)
            .Select(font => ((int)font.Key).ToString(CultureInfo.InvariantCulture) + ":" + (font.Value.FontName ?? string.Empty) + ":" + font.Value.Data.Length.ToString(CultureInfo.InvariantCulture)));
    }

    private static bool HasEmbeddedFontSlot(PdfCore.PdfOptions pdfOptions, PdfCore.PdfStandardFont font) {
        return pdfOptions.EmbeddedFonts.ContainsKey(PdfCore.PdfStandardFontMapper.GetFontFamily(font));
    }

    private static void RegisterPresentationFonts(PdfCore.PdfOptions pdfOptions, PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options, bool preserveConfiguredFontSlots) {
        var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots = CreateRegisteredFontSlots(pdfOptions, preserveConfiguredFontSlots);
        double pageWidth = presentation.SlideSize.WidthPoints;
        double pageHeight = presentation.SlideSize.HeightPoints;
        IReadOnlyList<PptCore.PowerPointSlide> slides = presentation.Slides;
        for (int slideIndex = 0; slideIndex < slides.Count; slideIndex++) {
            PptCore.PowerPointSlide slide = slides[slideIndex];
            if (!options.IncludeHiddenSlides && slide.Hidden) {
                continue;
            }

            int slideNumber = slideIndex + 1;
            RegisterPresentationShapesFonts(slide.GetInheritedShapesForExport(), slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options);
            RegisterPresentationShapesFonts(slide.Shapes, slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options);
        }
    }

    private static void RegisterPresentationShapesFonts(IReadOnlyList<PptCore.PowerPointShape> shapes, int slideNumber, double pageWidth, double pageHeight, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            RegisterPresentationShapeFonts(shape, slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options);
        }
    }

    private static void RegisterPresentationShapeFonts(PptCore.PowerPointShape shape, int slideNumber, double pageWidth, double pageHeight, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PowerPointPdfSaveOptions options) {
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
            RegisterPresentationShapesFonts(groupShape.OwnerSlide.GetGroupChildren(groupShape), slideNumber, pageWidth, pageHeight, pdfOptions, registeredFamilies, registeredFontSlots, options);
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
        if (string.IsNullOrWhiteSpace(familyName)) {
            return;
        }

        string trimmedFamilyName = familyName!.Trim();
        if (!registeredFamilies.Add(trimmedFamilyName)) {
            return;
        }

        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(trimmedFamilyName, out PdfCore.PdfStandardFont standardFont)) {
            PdfCore.PdfStandardFont fontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(standardFont);
            if (registeredFontSlots.Add(fontFamily)) {
                pdfOptions.RegisterOfficeFontFamily(trimmedFamilyName, fontFamily);
            }
        }
    }

    private static HashSet<PdfCore.PdfStandardFont> CreateRegisteredFontSlots(PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots) {
        var registeredFontSlots = new HashSet<PdfCore.PdfStandardFont>();
        if (preserveConfiguredFontSlots) {
            AddRegisteredFontSlot(registeredFontSlots, pdfOptions.DefaultFont);
            AddRegisteredFontSlot(registeredFontSlots, pdfOptions.HeaderFont);
            AddRegisteredFontSlot(registeredFontSlots, pdfOptions.FooterFont);
        }

        foreach (PdfCore.PdfStandardFont embeddedFont in pdfOptions.EmbeddedFonts.Keys) {
            AddRegisteredFontSlot(registeredFontSlots, embeddedFont);
        }

        return registeredFontSlots;
    }

    private static void AddRegisteredFontSlot(HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PdfCore.PdfStandardFont font) {
        registeredFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(font));
    }
}
