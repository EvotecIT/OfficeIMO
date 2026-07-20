namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    internal static System.Collections.Generic.IReadOnlyList<PdfStandardFont> CollectGeneratedStandardFonts(IEnumerable<IPdfBlock> blocks, PdfOptions options) {
        return CollectGeneratedComplianceEvidence(document: null, blocks, options).StandardFonts;
    }

    internal static PdfGeneratedDocumentComplianceEvidence CollectGeneratedComplianceEvidence(PdfDocument? document, IEnumerable<IPdfBlock> blocks, PdfOptions options) {
        Guard.NotNull(blocks, nameof(blocks));
        Guard.NotNull(options, nameof(options));

        using System.IDisposable? generatedSectionLayout = document?.BeginGeneratedSectionLayout();
        using LayoutResult layout = LayoutBlocks(blocks, options);
        return CollectGeneratedComplianceEvidence(layout, options);
    }

    private static PdfGeneratedDocumentComplianceEvidence CollectGeneratedComplianceEvidence(LayoutResult layout, PdfOptions options) {
        Guard.NotNull(layout, nameof(layout));
        Guard.NotNull(options, nameof(options));

        var fonts = new System.Collections.Generic.HashSet<PdfStandardFont>();
        var fontUsages = new System.Collections.Generic.List<PdfGeneratedFontComplianceEvidence>();
        var images = new System.Collections.Generic.List<PdfGeneratedImageAccessibilityEvidence>();
        var drawings = new System.Collections.Generic.List<PdfGeneratedDrawingAccessibilityEvidence>();
        var forms = new System.Collections.Generic.List<PdfGeneratedFormAccessibilityEvidence>();
        System.Collections.Generic.IReadOnlyList<PageNumberInfo> pageNumberInfos = BuildPageNumberInfos(layout.Pages);

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            LayoutResult.Page page = layout.Pages[pageIndex];
            PdfOptions pageOptions = page.Options ?? options;
            PdfStandardFont normalFont = ChooseNormal(pageOptions.DefaultFont);

            fonts.Add(normalFont);
            AddGeneratedFontUsage(fontUsages, normalFont, pageOptions);
            if (page.UsedBold) {
                PdfStandardFont boldFont = ChooseBold(normalFont);
                fonts.Add(boldFont);
                AddGeneratedFontUsage(fontUsages, boldFont, pageOptions);
            }

            if (page.UsedItalic) {
                PdfStandardFont italicFont = ChooseItalic(normalFont);
                fonts.Add(italicFont);
                AddGeneratedFontUsage(fontUsages, italicFont, pageOptions);
            }

            if (page.UsedBoldItalic) {
                PdfStandardFont boldItalicFont = ChooseBoldItalic(normalFont);
                fonts.Add(boldItalicFont);
                AddGeneratedFontUsage(fontUsages, boldItalicFont, pageOptions);
            }

            foreach (PdfStandardFont usedFont in page.UsedFonts) {
                fonts.Add(usedFont);
                AddGeneratedFontUsage(fontUsages, usedFont, pageOptions);
            }

            foreach (PdfNamedFontFace usedFont in page.UsedNamedFonts) {
                AddGeneratedFontUsage(fontUsages, usedFont, pageOptions);
            }

            int variantPageNumber = pageNumberInfos[pageIndex].VariantPageNumber;
            PdfTextWatermark? textWatermark = pageOptions.GetTextWatermarkForPage(variantPageNumber);
            if (textWatermark != null && textWatermark.Opacity > 0D) {
                PdfStandardFont watermarkFont = GetTextWatermarkFont(textWatermark);
                fonts.Add(watermarkFont);
                AddGeneratedFontUsage(fontUsages, watermarkFont, pageOptions);
            }

            if (pageOptions.HasHeaderTextContentForPage(variantPageNumber)) {
                if (TryResolvePageTextNamedFont(pageOptions, pageOptions.HeaderFontFamily, pageOptions.HeaderFont, out PdfNamedFontFace headerNamedFont)) {
                    AddGeneratedFontUsage(fontUsages, headerNamedFont, pageOptions);
                } else {
                    fonts.Add(pageOptions.HeaderFont);
                    AddGeneratedFontUsage(fontUsages, pageOptions.HeaderFont, pageOptions);
                }
            }

            if (pageOptions.HasFooterTextContentForPage(variantPageNumber)) {
                if (TryResolvePageTextNamedFont(pageOptions, pageOptions.FooterFontFamily, pageOptions.FooterFont, out PdfNamedFontFace footerNamedFont)) {
                    AddGeneratedFontUsage(fontUsages, footerNamedFont, pageOptions);
                } else {
                    fonts.Add(pageOptions.FooterFont);
                    AddGeneratedFontUsage(fontUsages, pageOptions.FooterFont, pageOptions);
                }
            }

            if (page.FormFields.Count > 0) {
                fonts.Add(PdfStandardFont.Helvetica);
                AddGeneratedFontUsage(fontUsages, PdfStandardFont.Helvetica, pageOptions);
                AddGeneratedFontUsage(fontUsages, PdfStandardFont.Helvetica, options);
                foreach (FormFieldAnnotation formField in page.FormFields) {
                    forms.Add(new PdfGeneratedFormAccessibilityEvidence(
                        formField.Name,
                        GetFormFieldWidgetCount(formField),
                        !string.IsNullOrWhiteSpace(formField.Style.AlternateName)));
                }
            }

            foreach (PageImage image in page.Images) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(!string.IsNullOrWhiteSpace(image.AlternativeText), image.IsDecorativeArtifact));
            }

            drawings.AddRange(page.Drawings);

            foreach (PdfHeaderFooterImage image in pageOptions.GetHeaderImagesForPage(variantPageNumber)) {
                bool hasAlternativeText = !string.IsNullOrWhiteSpace(image.AlternativeText);
                images.Add(new PdfGeneratedImageAccessibilityEvidence(hasAlternativeText, isDecorativeArtifact: !hasAlternativeText));
            }

            foreach (PdfHeaderFooterImage image in pageOptions.GetFooterImagesForPage(variantPageNumber)) {
                bool hasAlternativeText = !string.IsNullOrWhiteSpace(image.AlternativeText);
                images.Add(new PdfGeneratedImageAccessibilityEvidence(hasAlternativeText, isDecorativeArtifact: !hasAlternativeText));
            }

            PdfPageBackgroundImage? pageBackgroundImage = pageOptions.PageBackgroundImageSnapshot;
            if (pageBackgroundImage != null && pageBackgroundImage.Opacity > 0D) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(hasAlternativeText: false, isDecorativeArtifact: true));
            }

            PdfImageWatermark? imageWatermark = pageOptions.GetImageWatermarkForPage(variantPageNumber);
            if (imageWatermark != null && imageWatermark.Opacity > 0D) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(hasAlternativeText: false, isDecorativeArtifact: true));
            }
        }

        PdfStandardFont[] fontSnapshot = fonts
            .OrderBy(font => (int)font)
            .ToArray();
        return new PdfGeneratedDocumentComplianceEvidence(fontSnapshot, fontUsages.ToArray(), images.ToArray(), drawings.ToArray(), forms.ToArray());
    }

    private static void AddGeneratedFontUsage(System.Collections.Generic.List<PdfGeneratedFontComplianceEvidence> usages, PdfStandardFont font, PdfOptions options) {
        for (int i = 0; i < usages.Count; i++) {
            PdfGeneratedFontComplianceEvidence usage = usages[i];
            if (usage.StandardFont == font && !usage.NamedFont.HasValue && object.ReferenceEquals(usage.Options, options)) {
                return;
            }
        }

        usages.Add(new PdfGeneratedFontComplianceEvidence(font, options));
    }

    private static void AddGeneratedFontUsage(System.Collections.Generic.List<PdfGeneratedFontComplianceEvidence> usages, PdfNamedFontFace font, PdfOptions options) {
        for (int i = 0; i < usages.Count; i++) {
            PdfGeneratedFontComplianceEvidence usage = usages[i];
            if (usage.NamedFont == font && !usage.StandardFont.HasValue && object.ReferenceEquals(usage.Options, options)) {
                return;
            }
        }

        usages.Add(new PdfGeneratedFontComplianceEvidence(font, options));
    }

    private static int GetFormFieldWidgetCount(FormFieldAnnotation formField) {
        if (formField.Kind == FormFieldAnnotationKind.RadioButtonGroup) {
            return formField.Options.Count;
        }

        return 1;
    }
}
