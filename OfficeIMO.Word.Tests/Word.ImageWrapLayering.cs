using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {
        public static TheoryData<WordImageTextWrapping, WordImageTextWrapping?> FloatingWrapRoutes() {
            var data = new TheoryData<WordImageTextWrapping, WordImageTextWrapping?>();
            WordImageTextWrapping[] floating = {
                WordImageTextWrapping.Square, WordImageTextWrapping.Tight, WordImageTextWrapping.Through,
                WordImageTextWrapping.TopAndBottom, WordImageTextWrapping.BehindText, WordImageTextWrapping.InFrontOfText
            };
            foreach (WordImageTextWrapping wrap in floating) {
                data.Add(wrap, null);
                data.Add(wrap, WordImageTextWrapping.InLineWithText);
                data.Add(wrap, WordImageTextWrapping.BehindText);
                data.Add(wrap, WordImageTextWrapping.InFrontOfText);
                data.Add(wrap, WordImageTextWrapping.Square);
            }
            return data;
        }

        // Word saves behindDoc="1" only for behind-text pictures and keeps text wrapping
        // and effect extents in schema order; every route that sets a wrap must do the same.
        [Theory]
        [MemberData(nameof(FloatingWrapRoutes))]
        public void WordImage_BehindDocFollowsWrapForEveryRoute(WordImageTextWrapping wrap, WordImageTextWrapping? startingWrap) {
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            using WordDocument document = WordDocument.Create();
            WordImage image = startingWrap is { } start
                ? document.AddParagraph("Text").AddImage(imagePath, OfficeIMO.OfficeLength.Pixels(50), OfficeIMO.OfficeLength.Pixels(50),
                    options: new WordImageOptions { Wrapping = start })
                : document.AddParagraph("Text").AddImage(imagePath, OfficeIMO.OfficeLength.Pixels(50), OfficeIMO.OfficeLength.Pixels(50),
                    options: new WordImageOptions { Wrapping = wrap });
            if (startingWrap != null) image.WrapText = wrap;
            WordTextBox textBox = document.AddTextBox("Box", startingWrap is { } boxStart && boxStart != WordImageTextWrapping.InLineWithText ? boxStart : wrap);
            textBox.WrapText = wrap;

            using var stream = new MemoryStream(document.ToBytes());
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Anchor[] anchors = package.MainDocumentPart!.Document.Descendants<Anchor>().ToArray();
            Assert.Equal(2, anchors.Length);
            foreach (Anchor anchor in anchors) {
                Assert.Equal(wrap == WordImageTextWrapping.BehindText, anchor.BehindDoc?.Value);
                Assert.Single(anchor.ChildElements, static element =>
                    element is WrapNone or WrapSquare or WrapTight or WrapThrough or WrapTopBottom);
            }
            EffectExtent extent = Assert.IsType<EffectExtent>(anchors[0].EffectExtent);
            Assert.Equal(new long?[] { 0L, 0L, 0L, 0L },
                new[] { extent.LeftEdge?.Value, extent.TopEdge?.Value, extent.RightEdge?.Value, extent.BottomEdge?.Value });
        }
    }
}
