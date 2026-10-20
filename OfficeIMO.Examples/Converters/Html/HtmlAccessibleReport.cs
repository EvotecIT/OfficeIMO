using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Html;

internal static partial class Html {
    /// <summary>Renders a tagged report using the same embedded fonts for layout and PDF text.</summary>
    public static HtmlPdfRenderRequestResult CreateAccessibleReport(string html, byte[] regularFont,
        byte[] boldFont, PdfComplianceProfile profile = PdfComplianceProfile.PdfUa1) {
        var options = new HtmlToPdfOptions {
            DefaultFontFamily = "Helvetica",
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new PdfOptions().UsePdfUa(profile, "en")
                .EmbedStandardFont(PdfStandardFont.Helvetica, regularFont, "Carlito")
                .EmbedStandardFont(PdfStandardFont.HelveticaBold, boldFont, "Carlito Bold")
        };
        options.Fonts.Add("Helvetica", regularFont).Add("Helvetica", boldFont, OfficeFontStyle.Bold);
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        HtmlPdfAccessibilityValidationResult validation = HtmlPdfAccessibilityValidator.Validate(result);
        if (!validation.IsValid)
            throw new InvalidDataException(string.Join(Environment.NewLine, validation.Issues.Select(issue => issue.Message)));
        result.Output.RequireNoLoss();
        return result;
    }

    public static void Example_HtmlAccessibleReport(string folderPath, bool openPdf) {
        string html = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "Reports", "accessibility-review.html"));
        byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts", "Carlito-Regular.ttf"));
        byte[] bold = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts", "Carlito-Bold.ttf"));
        foreach (PdfComplianceProfile profile in new[] { PdfComplianceProfile.PdfUa1, PdfComplianceProfile.PdfUa2 }) {
            HtmlPdfRenderRequestResult result = CreateAccessibleReport(html, regular, bold, profile);
            string path = Path.Combine(folderPath, "HtmlAccessibleReport-" + profile + ".pdf");
            File.WriteAllBytes(path, result.ToBytes());
            Console.WriteLine("Created " + path + "; validate these exact bytes independently and complete the reading review.");
            if (openPdf)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
    }
}
