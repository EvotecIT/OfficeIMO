using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FillFields_UsesCurrentStreamPositions() {
        byte[] source = BuildHierarchicalFormPdf();
        using var input = new MemoryStream();
        input.WriteByte(123);
        input.Write(source, 0, source.Length);
        input.Position = 1;
        using var output = new MemoryStream();

        PdfFormFiller.FillFields(input, output, new Dictionary<string, string> {
            ["Person.Name"] = "Stream value"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(output.ToArray());
        Assert.Equal("Stream value", info.FormFields[0].Value);
    }

    [Fact]
    public void FillFields_PathHelpersWriteFilledPdf() {
        string inputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-fill-input-" + Guid.NewGuid().ToString("N") + ".pdf");
        string outputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-fill-output-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(inputPath, BuildHierarchicalFormPdf());

            PdfFormFiller.FillFields(inputPath, outputPath, new Dictionary<string, string> {
                ["Person.Name"] = "Path value"
            });

            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal("Path value", info.FormFields[0].Value);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void FlattenFields_PathHelpersWriteFlattenedPdf() {
        string inputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-flatten-input-" + Guid.NewGuid().ToString("N") + ".pdf");
        string outputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-flatten-output-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Path flatten"
            });
            File.WriteAllBytes(inputPath, filled);

            PdfFormFiller.FlattenFields(inputPath, outputPath);

            byte[] flattened = File.ReadAllBytes(outputPath);
            Assert.False(PdfInspector.Inspect(flattened).HasForms);
            Assert.Contains("<5061746820666C617474656E> Tj", Encoding.ASCII.GetString(flattened));
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void FormPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string fillInputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-fill-stream-input-" + Guid.NewGuid().ToString("N") + ".pdf");
        string flattenInputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-flatten-stream-input-" + Guid.NewGuid().ToString("N") + ".pdf");
        string fillFlattenInputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-fill-flatten-stream-input-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(fillInputPath, BuildHierarchicalFormPdf());
            byte[] filledForFlatten = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
                ["Name"] = "Path stream flatten"
            });
            File.WriteAllBytes(flattenInputPath, filledForFlatten);
            File.WriteAllBytes(fillFlattenInputPath, BuildTextWidgetFormPdf());

            using var fillOutput = new MemoryStream();
            fillOutput.WriteByte(17);
            PdfFormFiller.FillFields(fillInputPath, fillOutput, new Dictionary<string, string> {
                ["Person.Name"] = "Path stream fill"
            });
            byte[] fillBytes = SliceAfterPrefix(fillOutput, 1);
            Assert.Equal(17, fillOutput.ToArray()[0]);
            Assert.Equal("Path stream fill", PdfInspector.Inspect(fillBytes).FormFields[0].Value);

            using var flattenOutput = new MemoryStream();
            flattenOutput.WriteByte(23);
            PdfFormFiller.FlattenFields(flattenInputPath, flattenOutput);
            byte[] flattenBytes = SliceAfterPrefix(flattenOutput, 1);
            Assert.Equal(23, flattenOutput.ToArray()[0]);
            Assert.False(PdfInspector.Inspect(flattenBytes).HasForms);
            Assert.Contains("<506174682073747265616D20666C617474656E> Tj", Encoding.ASCII.GetString(flattenBytes));

            using var fillFlattenOutput = new MemoryStream();
            fillFlattenOutput.WriteByte(29);
            PdfFormFiller.FillAndFlattenFields(fillFlattenInputPath, fillFlattenOutput, new Dictionary<string, string> {
                ["Name"] = "Path stream single pass"
            });
            byte[] fillFlattenBytes = SliceAfterPrefix(fillFlattenOutput, 1);
            Assert.Equal(29, fillFlattenOutput.ToArray()[0]);
            Assert.False(PdfInspector.Inspect(fillFlattenBytes).HasForms);
            Assert.Contains("<506174682073747265616D2073696E676C652070617373> Tj", Encoding.ASCII.GetString(fillFlattenBytes));
        } finally {
            if (File.Exists(fillInputPath)) File.Delete(fillInputPath);
            if (File.Exists(flattenInputPath)) File.Delete(flattenInputPath);
            if (File.Exists(fillFlattenInputPath)) File.Delete(fillFlattenInputPath);
        }
    }
}
