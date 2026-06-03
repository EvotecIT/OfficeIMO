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
    public void VectorRectangle_RendersFillAndStrokeOperators() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Rectangle(
                width: 100,
                height: 36,
                strokeColor: PdfColor.FromRgb(26, 51, 77),
                strokeWidth: 2.5,
                fillColor: PdfColor.FromRgb(204, 179, 153),
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.8 0.702 0.6 rg", content);
        Assert.Contains("0.102 0.2 0.302 RG", content);
        Assert.Contains("2.5 w", content);
        Assert.Contains("70 124 100 36 re B", content);
    }


}
