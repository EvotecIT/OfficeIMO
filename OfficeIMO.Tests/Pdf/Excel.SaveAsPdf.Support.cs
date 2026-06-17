using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            if (text.IndexOf(word, StringComparison.Ordinal) >= 0) {
                return ordered[0].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindFirstLetterStartX(UglyToad.PdfPig.Content.Page page, string letter) {
        double x = page.Letters
            .Where(pdfLetter => string.Equals(pdfLetter.Value, letter, StringComparison.Ordinal))
            .Select(pdfLetter => pdfLetter.StartBaseLine.X)
            .DefaultIfEmpty(double.NaN)
            .First();

        if (double.IsNaN(x)) {
            throw new InvalidOperationException("Could not find letter '" + letter + "' in rendered PDF text.");
        }

        return x;
    }

    private static byte[] CreateMinimalRgbPng() => Pdf.PdfPngTestImages.CreateRgbPng(255, 0, 0);

    private static string NormalizePdfTextSpaces(string text) =>
        text.Replace('\u00A0', ' ').Replace('\u202F', ' ');
}
