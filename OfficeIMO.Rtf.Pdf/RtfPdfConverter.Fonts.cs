using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static IReadOnlyDictionary<int, PdfCore.PdfStandardFont> ConfigureDocumentFonts(
        RtfDocument document,
        PdfCore.PdfOptions pdfOptions,
        RtfPdfSaveOptions options) {
        bool allowSystemFontEmbedding = options.ResourcePolicy.AllowSystemFontEmbedding;
        var fontSlots = new Dictionary<int, PdfCore.PdfStandardFont>();
        var familySlots = new Dictionary<string, PdfCore.PdfStandardFont>(StringComparer.OrdinalIgnoreCase);
        HashSet<int> referencedFontIds = CollectReferencedFontIds(document);
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(includeDocumentFontSlots: false);
        if (pdfOptions.HasExplicitDefaultFont) PdfCore.PdfOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.DefaultFont);
        if (pdfOptions.HasExplicitHeaderFont) PdfCore.PdfOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.HeaderFont);
        if (pdfOptions.HasExplicitFooterFont) PdfCore.PdfOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.FooterFont);

        RtfFont? defaultFont = document.Settings.DefaultFontId.HasValue
            ? document.Fonts.FirstOrDefault(font => font.Id == document.Settings.DefaultFontId.Value)
            : null;
        if (defaultFont != null) {
            bool preserveConfiguredDefaultFont = pdfOptions.HasExplicitDefaultFont ||
                pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont);
            if (preserveConfiguredDefaultFont) {
                PdfCore.PdfStandardFont configuredDefaultSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(pdfOptions.DefaultFont);
                registeredFontSlots.Add(configuredDefaultSlot);
                familySlots[defaultFont.Name] = configuredDefaultSlot;
                fontSlots[defaultFont.Id] = configuredDefaultSlot;
            } else if (TryRegisterDocumentFont(
                defaultFont.Name,
                pdfOptions,
                registeredFontSlots,
                allowSystemFontEmbedding,
                options,
                out PdfCore.PdfStandardFont defaultSlot)) {
                pdfOptions.DefaultFont = defaultSlot;
                familySlots[defaultFont.Name] = defaultSlot;
                fontSlots[defaultFont.Id] = defaultSlot;
            }
        }

        foreach (RtfFont font in document.Fonts) {
            if (!referencedFontIds.Contains(font.Id)) {
                continue;
            }

            if (familySlots.TryGetValue(font.Name, out PdfCore.PdfStandardFont existingSlot)) {
                fontSlots[font.Id] = existingSlot;
                continue;
            }

            if (TryRegisterDocumentFont(
                font.Name,
                pdfOptions,
                registeredFontSlots,
                allowSystemFontEmbedding,
                options,
                out PdfCore.PdfStandardFont fontSlot)) {
                familySlots[font.Name] = fontSlot;
                fontSlots[font.Id] = fontSlot;
            }
        }

        return fontSlots;
    }

    private static HashSet<int> CollectReferencedFontIds(RtfDocument document) {
        var fontIds = new HashSet<int>();
        if (document.Settings.DefaultFontId.HasValue) {
            fontIds.Add(document.Settings.DefaultFontId.Value);
        }

        if (document.Sections.Count > 0) {
            foreach (RtfSection section in document.Sections) {
                CollectReferencedFontIds(section.Blocks, fontIds);
            }
        } else {
            CollectReferencedFontIds(document.Blocks, fontIds);
        }

        foreach (RtfNote note in document.Notes) {
            foreach (RtfParagraph paragraph in note.Paragraphs) {
                CollectReferencedFontIds(paragraph, fontIds);
            }
        }

        return fontIds;
    }

    private static void CollectReferencedFontIds(IEnumerable<IRtfBlock> blocks, HashSet<int> fontIds) {
        foreach (IRtfBlock block in blocks) {
            if (block is RtfParagraph paragraph) {
                CollectReferencedFontIds(paragraph, fontIds);
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        CollectReferencedFontIds(cell.Blocks, fontIds);
                    }
                }
            }
        }
    }

    private static void CollectReferencedFontIds(RtfParagraph paragraph, HashSet<int> fontIds) {
        foreach (RtfRun run in paragraph.Runs) {
            if (run.FontId.HasValue) fontIds.Add(run.FontId.Value);
        }

        foreach (IRtfInline inline in paragraph.Inlines) {
            if (inline is RtfField field) {
                CollectReferencedFontIds(field.Result, fontIds);
            }
        }
    }

    private static bool TryRegisterDocumentFont(
        string familyName,
        PdfCore.PdfOptions pdfOptions,
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
        bool allowSystemFontEmbedding,
        RtfPdfSaveOptions options,
        out PdfCore.PdfStandardFont fontSlot) {
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont mappedFont)) {
            PdfCore.PdfStandardFont mappedFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont);
            if (!registeredFontSlots.Contains(mappedFamily)) {
                pdfOptions.RegisterOfficeFontFamily(familyName, mappedFamily, allowSystemFontEmbedding);
                registeredFontSlots.Add(mappedFamily);
                fontSlot = mappedFamily;
                return true;
            }

            if (pdfOptions.EmbeddedFontFamilySlotMatches(mappedFamily, familyName)) {
                fontSlot = mappedFamily;
                return true;
            }

            if (allowSystemFontEmbedding &&
                PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont distinctSlot) &&
                PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? distinctFamily) &&
                distinctFamily != null) {
                pdfOptions.RegisterFontFamily(distinctSlot, distinctFamily);
                registeredFontSlots.Add(distinctSlot);
                fontSlot = distinctSlot;
                return true;
            }

            if (pdfOptions.HasEmbeddedStandardFontFamily(mappedFamily)) {
                ReportFontSlotExhaustion(options, familyName, mappedFamily, pdfOptions.GetEmbeddedFontFamilyName(mappedFamily));
                fontSlot = PdfCore.PdfStandardFont.Helvetica;
                return false;
            }

            fontSlot = mappedFamily;
            return true;
        }

        if (allowSystemFontEmbedding) {
            if (!PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont availableSlot)) {
                ReportFontSlotExhaustion(options, familyName, null, null);
            } else if (PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? embeddedFamily) &&
                       embeddedFamily != null) {
                pdfOptions.RegisterFontFamily(availableSlot, embeddedFamily);
                registeredFontSlots.Add(availableSlot);
                fontSlot = availableSlot;
                return true;
            }
        }

        fontSlot = PdfCore.PdfStandardFont.Helvetica;
        return false;
    }

    private static void ReportFontSlotExhaustion(
        RtfPdfSaveOptions options,
        string familyName,
        PdfCore.PdfStandardFont? fallbackSlot,
        string? occupyingFontFamily) {
        var details = new Dictionary<string, string> {
            ["fontFamily"] = familyName
        };
        if (fallbackSlot.HasValue) details["fallbackSlot"] = fallbackSlot.Value.ToString();
        if (!string.IsNullOrWhiteSpace(occupyingFontFamily)) details["occupyingFontFamily"] = occupyingFontFamily!;
        AddConversionWarning(
            options,
            "FontFamilySlotExhausted",
            "Font/" + familyName,
            "The RTF font family could not receive a distinct embedded PDF family slot; affected runs use the configured document default instead.",
            RtfConversionAction.Substituted,
            details);
    }
}
