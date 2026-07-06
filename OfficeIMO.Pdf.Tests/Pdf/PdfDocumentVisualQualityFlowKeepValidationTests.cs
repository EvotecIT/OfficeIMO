using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Paragraph_KeepTogetherRejectsContentTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        string longText = string.Join(" ", Enumerable.Range(1, 180).Select(i => "paragraph" + i.ToString("000")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Paragraph(p => p.Text(longText), style: new PdfParagraphStyle {
                    KeepTogether = true
                })
                .ToBytes());

        Assert.Contains("Paragraph height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void List_KeepTogetherRejectsContentTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Bullets(Enumerable.Range(1, 14).Select(i => "Keep list item " + i.ToString("00")), style: new PdfListStyle {
                    KeepTogether = true
                })
                .ToBytes());

        Assert.Contains("List height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_KeepTogetherSplitsContentTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        string longText = string.Join(" ", Enumerable.Range(1, 180).Select(i => "panel" + i.ToString("000")));

        byte[] bytes = PdfDocument.Create(options)
            .PanelParagraph(p => p.Text(longText), new PanelStyle {
                KeepTogether = true,
                PaddingY = 8
            })
            .ToBytes();

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1);
    }



}
