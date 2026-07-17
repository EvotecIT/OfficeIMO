using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static IReadOnlyDictionary<int, PdfCore.PdfStandardFont> ConfigureDocumentFonts(
        RtfDocument document,
        PdfCore.PdfOptions pdfOptions,
        RtfPdfSaveOptions options,
        bool preserveConfiguredFontSlots) {
        bool allowSystemFontEmbedding = options.ResourcePolicy.AllowSystemFontEmbedding;
        var fontSlots = new Dictionary<int, PdfCore.PdfStandardFont>();
        var familySlots = new Dictionary<string, PdfCore.PdfStandardFont>(StringComparer.OrdinalIgnoreCase);
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots =
            pdfOptions.CreateRegisteredFontFamilySlots(preserveConfiguredFontSlots);

        RtfFont? defaultFont = document.Settings.DefaultFontId.HasValue
            ? document.Fonts.FirstOrDefault(font => font.Id == document.Settings.DefaultFontId.Value)
            : null;
        if (!preserveConfiguredFontSlots && defaultFont != null) {
            pdfOptions.UseOfficeFontFamily(defaultFont.Name, allowSystemFontEmbedding);
            bool hasMappedDefault = PdfCore.PdfStandardFontMapper.TryMapFontFamily(
                defaultFont.Name,
                out PdfCore.PdfStandardFont mappedDefault);
            if (hasMappedDefault || pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
                PdfCore.PdfStandardFont defaultSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(
                    hasMappedDefault ? mappedDefault : pdfOptions.DefaultFont);
                registeredFontSlots.Add(defaultSlot);
                familySlots[defaultFont.Name] = defaultSlot;
                fontSlots[defaultFont.Id] = defaultSlot;
            }
        }

        foreach (RtfFont font in document.Fonts) {
            if (familySlots.TryGetValue(font.Name, out PdfCore.PdfStandardFont existingSlot)) {
                fontSlots[font.Id] = existingSlot;
                continue;
            }

            if (TryRegisterDocumentFont(
                font.Name,
                pdfOptions,
                registeredFontSlots,
                allowSystemFontEmbedding,
                out PdfCore.PdfStandardFont fontSlot)) {
                familySlots[font.Name] = fontSlot;
                fontSlots[font.Id] = fontSlot;
            }
        }

        return fontSlots;
    }

    private static bool TryRegisterDocumentFont(
        string familyName,
        PdfCore.PdfOptions pdfOptions,
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
        bool allowSystemFontEmbedding,
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

            fontSlot = mappedFamily;
            return true;
        }

        if (allowSystemFontEmbedding &&
            PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont availableSlot) &&
            PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? embeddedFamily) &&
            embeddedFamily != null) {
            pdfOptions.RegisterFontFamily(availableSlot, embeddedFamily);
            registeredFontSlots.Add(availableSlot);
            fontSlot = availableSlot;
            return true;
        }

        fontSlot = PdfCore.PdfStandardFont.Helvetica;
        return false;
    }
}
