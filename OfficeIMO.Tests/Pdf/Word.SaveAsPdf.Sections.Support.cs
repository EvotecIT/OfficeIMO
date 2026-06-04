using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
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

    private static string NormalizePdfText(string text) =>
        string.Join(" ", text.Split(new[] { ' ', '\r', '\n', '\t', '\f' }, StringSplitOptions.RemoveEmptyEntries));

    private static void ReplaceFirstHeaderImagePartWithGif(string docPath) {
        byte[] gifBytes = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
        using WordprocessingDocument package = WordprocessingDocument.Open(docPath, true);
        HeaderPart headerPart = package.MainDocumentPart!.HeaderParts.First();
        ImagePart imagePart = headerPart.ImageParts.First();
        using var stream = new MemoryStream(gifBytes);
        imagePart.FeedData(stream);
    }
}
