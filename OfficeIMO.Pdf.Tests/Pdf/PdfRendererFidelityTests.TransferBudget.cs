using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Theory]
    [InlineData(2, false, true)]
    [InlineData(256, false, false)]
    [InlineData(300, true, true)]
    public void SoftMaskTransferTablesShareThePageWorkBudget(int maximum, bool calculator, bool exceeds) {
        byte[] group = Encoding.ASCII.GetBytes("0.25 g 0 0 100 100 re f");
        byte[] code = Encoding.ASCII.GetBytes("{ 1 exch sub }");
        var function = calculator
            ? ("<< /FunctionType 4 /Domain [0 1] /Range [0 1] /Length " + code.Length + " >>", code)
            : ("<< /FunctionType 2 /Domain [0 1] /C0 [1] /C1 [0] /N 1 >>", Array.Empty<byte>());
        var objects = new List<(string, byte[])> {
            ("<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /S /Transparency /CS /DeviceRGB >> /Length " + group.Length + " >>", group),
            function
        };
        byte[] pdf = BuildPdf("/G1 gs /G1 gs /G1 gs 1 0 0 rg 10 10 80 80 re f",
            "<< /ExtGState << /G1 << /SMask << /Type /Mask /S /Luminosity /G 5 0 R /TR 6 0 R >> >> >> >>", objects: objects);
        PdfReadPage page = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxFunctionShadingEvaluationWork = maximum }
        }).Pages[0];
        if (exceeds) {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => page.ExportImage(OfficeImageExportFormat.Png));
            Assert.Equal(PdfReadLimitKind.FunctionShadingEvaluationWork, error.Kind);
        } else {
            OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);
            AssertColor(PdfRendererQualitySupport.Decode(result.Bytes), 50, 50, 255, 64, 64, tolerance: 3);
        }
    }
}
