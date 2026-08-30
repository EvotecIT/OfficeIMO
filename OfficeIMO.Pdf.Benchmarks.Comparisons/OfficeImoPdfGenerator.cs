using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class OfficeImoPdfGenerator {
    internal static byte[] Generate(PdfBenchmarkScenario scenario) {
        var options = new PdfOptions {
            CompressContentStreams = true,
            DefaultFontSize = 9,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly
        };
        options.RegisterFontFamily(
            PdfStandardFont.Helvetica,
            new PdfEmbeddedFontFamily("Carlito", PdfBenchmarkAssets.CarlitoRegular));

        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => {
            for (int page = 1; page <= scenario.PageCount; page++) {
                content.H1(scenario.PageTitle(page));
                for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                    string narrative = scenario.Narrative(page, paragraph);
                    content.Paragraph(builder => builder.Text(narrative));
                }
                content.Table(scenario.TableRows(page));
                if (page < scenario.PageCount) {
                    content.PageBreak();
                }
            }
        }), options).Meta(title: scenario.Name);

        return document.ToBytes();
    }

    internal static byte[] GenerateHtml(string html, PdfEmbeddedFontFamily? fontFamily = null) =>
        HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions {
            FontFamily = fontFamily
        });

}
