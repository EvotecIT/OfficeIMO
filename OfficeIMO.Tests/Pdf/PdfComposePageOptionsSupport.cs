using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfComposePageOptionsTests {
        private static string Normalize(string text) {
            return new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
        }

        private static void AssertMargins(PageMargins margins, double left, double top, double right, double bottom) {
            Assert.Equal(left, margins.Left, 6);
            Assert.Equal(top, margins.Top, 6);
            Assert.Equal(right, margins.Right, 6);
            Assert.Equal(bottom, margins.Bottom, 6);
        }

        private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
            var lines = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => System.Math.Round(letter.StartBaseLine.Y, 1));

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
                .GroupBy(letter => System.Math.Round(letter.StartBaseLine.Y, 1));

            foreach (var line in lines) {
                var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
                string text = string.Concat(ordered.Select(letter => letter.Value));
                int index = text.IndexOf(word, StringComparison.Ordinal);
                if (index >= 0) {
                    return ordered[index].StartBaseLine.Y;
                }
            }

            throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
        }

        private static byte[] CreateMinimalRgbPng() {
            return new byte[] {
                137, 80, 78, 71, 13, 10, 26, 10,
                0, 0, 0, 13,
                73, 72, 68, 82,
                0, 0, 0, 1,
                0, 0, 0, 1,
                8, 2, 0, 0, 0,
                0, 0, 0, 0,
                0, 0, 0, 12,
                73, 68, 65, 84,
                0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
                0, 0, 0, 0,
                0, 0, 0, 0,
                73, 69, 78, 68,
                0, 0, 0, 0
            };
        }
    }
}
