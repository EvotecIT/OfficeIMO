using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void WriterFontSelection_NormalizesEveryStandardFontVariantToItsFamily() {
            Assert.Equal(PdfStandardFont.Helvetica, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.HelveticaOblique));
            Assert.Equal(PdfStandardFont.Helvetica, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.HelveticaBoldOblique));
            Assert.Equal(PdfStandardFont.TimesRoman, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.TimesItalic));
            Assert.Equal(PdfStandardFont.TimesRoman, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.TimesBoldItalic));
            Assert.Equal(PdfStandardFont.Courier, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.CourierOblique));
            Assert.Equal(PdfStandardFont.Courier, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.CourierBoldOblique));

            Assert.Equal(PdfStandardFont.TimesBold, InvokePrivateFontMethod<PdfStandardFont>("ChooseBold", PdfStandardFont.TimesBoldItalic));
            Assert.Equal(PdfStandardFont.CourierOblique, InvokePrivateFontMethod<PdfStandardFont>("ChooseItalic", PdfStandardFont.CourierBoldOblique));
            Assert.Equal(PdfStandardFont.HelveticaBoldOblique, InvokePrivateFontMethod<PdfStandardFont>("ChooseBoldItalic", PdfStandardFont.HelveticaBoldOblique));
        }

        [Fact]
        public void WriterFontSelection_RejectsInvalidFontValuesInsteadOfFallingBack() {
            var chooseException = InvokePrivateFontMethodExpectingFailure("ChooseNormal", (PdfStandardFont)99);
            Assert.IsType<ArgumentOutOfRangeException>(chooseException.InnerException);

            var glyphException = InvokePrivateFontMethodExpectingFailure("GlyphWidthEmFor", (PdfStandardFont)99);
            Assert.IsType<ArgumentOutOfRangeException>(glyphException.InnerException);

            var spaceException = InvokePrivateFontMethodExpectingFailure("SpaceWidthEmFor", (PdfStandardFont)99);
            Assert.IsType<ArgumentOutOfRangeException>(spaceException.InnerException);

            var ascenderException = InvokePrivateFontMethodExpectingFailure("GetAscender", (PdfStandardFont)99, 12.0);
            Assert.IsType<ArgumentOutOfRangeException>(ascenderException.InnerException);

            var descenderException = InvokePrivateFontMethodExpectingFailure("GetDescender", (PdfStandardFont)99, 12.0);
            Assert.IsType<ArgumentOutOfRangeException>(descenderException.InnerException);
        }

        [Fact]
        public void WriterFontSelection_UsesRequestedFontFamilyForGeneratedPdfResources() {
            var pdf = PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.TimesItalic,
                    FooterFont = PdfStandardFont.TimesRoman
                })
                .Paragraph(p => p.Text("Times family should stay Times."))
                .ToBytes();

            string content = System.Text.Encoding.ASCII.GetString(pdf);

            Assert.Contains("/BaseFont /Times-Roman", content);
            Assert.DoesNotContain("/BaseFont /Courier", content);
        }

        [Fact]
        public void WriterFontSelection_UsesRunFontForGeneratedPdfResources() {
            var pdf = PdfDocument.Create()
                .Paragraph(p => p
                    .Text("Default ")
                    .Font(PdfStandardFont.Courier)
                    .Text("Code")
                    .ResetFont()
                    .Text(" Default"))
                .ToBytes();

            string content = System.Text.Encoding.ASCII.GetString(pdf);

            Assert.Contains("/BaseFont /Helvetica", content);
            Assert.Contains("/BaseFont /Courier", content);
        }
    }
}
