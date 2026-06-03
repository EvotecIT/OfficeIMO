namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    internal static System.Collections.Generic.IReadOnlyList<PdfStandardFont> CollectGeneratedStandardFonts(IEnumerable<IPdfBlock> blocks, PdfOptions options) {
        return CollectGeneratedComplianceEvidence(blocks, options).StandardFonts;
    }

    internal static PdfGeneratedDocumentComplianceEvidence CollectGeneratedComplianceEvidence(IEnumerable<IPdfBlock> blocks, PdfOptions options) {
        Guard.NotNull(blocks, nameof(blocks));
        Guard.NotNull(options, nameof(options));

        LayoutResult layout = LayoutBlocks(blocks, options);
        var fonts = new System.Collections.Generic.HashSet<PdfStandardFont>();
        var images = new System.Collections.Generic.List<PdfGeneratedImageAccessibilityEvidence>();
        var drawings = new System.Collections.Generic.List<PdfGeneratedDrawingAccessibilityEvidence>();
        var forms = new System.Collections.Generic.List<PdfGeneratedFormAccessibilityEvidence>();
        System.Collections.Generic.IReadOnlyList<PageNumberInfo> pageNumberInfos = BuildPageNumberInfos(layout.Pages);

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            LayoutResult.Page page = layout.Pages[pageIndex];
            PdfOptions pageOptions = page.Options ?? options;
            PdfStandardFont normalFont = ChooseNormal(pageOptions.DefaultFont);

            fonts.Add(normalFont);
            if (page.UsedBold) {
                fonts.Add(ChooseBold(normalFont));
            }

            if (page.UsedItalic) {
                fonts.Add(ChooseItalic(normalFont));
            }

            if (page.UsedBoldItalic) {
                fonts.Add(ChooseBoldItalic(normalFont));
            }

            foreach (PdfStandardFont usedFont in page.UsedFonts) {
                fonts.Add(usedFont);
            }

            PdfTextWatermark? textWatermark = pageOptions.TextWatermarkSnapshot;
            if (textWatermark != null && textWatermark.Opacity > 0D) {
                fonts.Add(GetTextWatermarkFont(textWatermark));
            }

            int variantPageNumber = pageNumberInfos[pageIndex].VariantPageNumber;
            if (pageOptions.HasHeaderTextContentForPage(variantPageNumber)) {
                fonts.Add(pageOptions.HeaderFont);
            }

            if (pageOptions.HasFooterTextContentForPage(variantPageNumber)) {
                fonts.Add(pageOptions.FooterFont);
            }

            if (page.FormFields.Count > 0) {
                fonts.Add(PdfStandardFont.Helvetica);
                foreach (FormFieldAnnotation formField in page.FormFields) {
                    forms.Add(new PdfGeneratedFormAccessibilityEvidence(
                        formField.Name,
                        GetFormFieldWidgetCount(formField),
                        !string.IsNullOrWhiteSpace(formField.Style.AlternateName)));
                }
            }

            foreach (PageImage image in page.Images) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(!string.IsNullOrWhiteSpace(image.AlternativeText), image.IsBackgroundDecoration));
            }

            drawings.AddRange(page.Drawings);

            foreach (PdfHeaderFooterImage image in pageOptions.GetHeaderImagesForPage(variantPageNumber)) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(!string.IsNullOrWhiteSpace(image.AlternativeText), isDecorativeArtifact: false));
            }

            foreach (PdfHeaderFooterImage image in pageOptions.GetFooterImagesForPage(variantPageNumber)) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(!string.IsNullOrWhiteSpace(image.AlternativeText), isDecorativeArtifact: false));
            }

            PdfPageBackgroundImage? pageBackgroundImage = pageOptions.PageBackgroundImageSnapshot;
            if (pageBackgroundImage != null && pageBackgroundImage.Opacity > 0D) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(hasAlternativeText: false, isDecorativeArtifact: true));
            }

            PdfImageWatermark? imageWatermark = pageOptions.ImageWatermarkSnapshot;
            if (imageWatermark != null && imageWatermark.Opacity > 0D) {
                images.Add(new PdfGeneratedImageAccessibilityEvidence(hasAlternativeText: false, isDecorativeArtifact: true));
            }
        }

        PdfStandardFont[] fontSnapshot = fonts
            .OrderBy(font => (int)font)
            .ToArray();
        return new PdfGeneratedDocumentComplianceEvidence(fontSnapshot, images.ToArray(), drawings.ToArray(), forms.ToArray());
    }

    private static int GetFormFieldWidgetCount(FormFieldAnnotation formField) {
        if (formField.Kind == FormFieldAnnotationKind.RadioButtonGroup) {
            return formField.Options.Count;
        }

        return 1;
    }
}
