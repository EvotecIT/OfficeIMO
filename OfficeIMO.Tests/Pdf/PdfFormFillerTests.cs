using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormFillerTests {
    [Fact]
    public void FillFields_UpdatesSimpleTextAndButtonValues() {
        byte[] filled = PdfFormFiller.FillFields(BuildHierarchicalFormPdf(), new Dictionary<string, string> {
            ["Person.Name"] = "Evotec",
            ["AcceptTerms"] = "Off"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.True(info.HasReadableFormFields);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms" }, info.FormFieldNames);
        Assert.Equal("Evotec", info.FormFields[0].Value);
        Assert.Equal("Off", info.FormFields[1].Value);
        Assert.Contains("/NeedAppearances true", Encoding.ASCII.GetString(filled));
        Assert.False(PdfInspector.Preflight(filled).CanRewrite);
    }

    [Fact]
    public void FillFields_GeneratesSimpleTextWidgetAppearance() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Visible value"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Equal("Visible value", info.FormFields[0].Value);
        Assert.Contains("/Subtype /Form", output);
        Assert.Contains("/AP << /N", output);
        Assert.Contains("/Helv", output);
        Assert.Contains("(Visible value) Tj", output);
    }

    [Fact]
    public void FillFields_GeneratesSimpleButtonWidgetAppearances() {
        byte[] filled = PdfFormFiller.FillFields(BuildCheckboxWidgetWithoutAppearancePdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Equal("Yes", info.FormFields[0].Value);
        Assert.Contains("/AS /Yes", output);
        Assert.Contains("/AP << /N <<", output);
        Assert.Contains("/Off", output);
        Assert.Contains("/Yes", output);
        Assert.Contains("1.25 w", output);
    }

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
    public void FlattenFields_PaintsTextWidgetAndRemovesFormAnnotations() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Flattened value"
        });

        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.False(info.HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("(Flattened value) Tj", output);
    }

    [Fact]
    public void FlattenFields_FlattensReferencedContentArraysBeforeAppendingAppearanceStream() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdfWithReferencedContentArray(), new Dictionary<string, string> {
            ["Name"] = "Flattened value"
        });

        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        var (objects, _) = PdfSyntax.ParseObjects(flattened);
        var page = Assert.IsType<PdfDictionary>(objects.Values.First(indirect =>
            indirect.Value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page").Value);
        var contents = Assert.IsType<PdfArray>(page.Items["Contents"]);

        Assert.True(contents.Items.Count >= 2);
        foreach (var item in contents.Items) {
            var reference = Assert.IsType<PdfReference>(item);
            Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
        }
    }

    [Fact]
    public void FillAndFlattenFields_UpdatesValueAndFlattensInOneCall() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Single pass"
        });

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.Contains("(Single pass) Tj", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
    }

    [Fact]
    public void FlattenFields_PaintsSelectedButtonWidgetAppearance() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildCheckboxWidgetFormPdf());

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("Checked appearance", GetFlattenedAppearanceStreamText(flattened));
    }

    [Fact]
    public void FillAndFlattenFields_UpdatesButtonStateBeforeFlattening() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildCheckboxWidgetFormPdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Off"
        });

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("Unchecked appearance", GetFlattenedAppearanceStreamText(flattened));
    }

    [Fact]
    public void FillAndFlattenFields_GeneratesMissingButtonAppearanceBeforeFlattening() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildCheckboxWidgetWithoutAppearancePdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string appearance = GetFlattenedAppearanceStreamText(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("1.25 w", appearance);
        Assert.Contains(" l S", appearance);
    }

    [Fact]
    public void FillAndFlattenFields_PaintsChoiceOptionDisplayText() {
        byte[] filled = PdfFormFiller.FillFields(BuildChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Country"] = "PL"
        });

        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.True(filledInfo.HasReadableFormFields);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);
        Assert.Equal("PL", filledField.Value);
        Assert.Equal("Poland", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("(Poland) Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);

        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo flattenedInfo = PdfInspector.Inspect(flattened);

        Assert.False(flattenedInfo.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("(Poland) Tj", GetFlattenedAppearanceStreamText(flattened), StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ChoiceDisplayTextStoresExportValueAndPaintsDisplayText() {
        byte[] filled = PdfFormFiller.FillFields(BuildChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Country"] = "United States"
        });

        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.True(filledInfo.HasReadableFormFields);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);
        Assert.Equal("US", filledField.Value);
        Assert.Equal("United States", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("(United States) Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void FlattenFields_PaintsMultiSelectChoiceOptionDisplayText() {
        byte[] source = BuildMultiSelectChoiceWidgetFormPdf();
        PdfDocumentInfo sourceInfo = PdfInspector.Inspect(source);

        PdfFormField sourceField = Assert.Single(sourceInfo.FormFields);
        Assert.Equal(new[] { "PL", "US" }, sourceField.Values);
        Assert.Equal(new[] { "Poland", "United States" }, sourceField.SelectedOptions.Select(option => option.DisplayText).ToArray());

        byte[] flattened = PdfFormFiller.FlattenFields(source);

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo flattenedInfo = PdfInspector.Inspect(flattened);

        Assert.False(flattenedInfo.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("(Poland, United States) Tj", GetFlattenedAppearanceStreamText(flattened), StringComparison.Ordinal);
    }

    [Fact]
    public void FlattenFields_GeneratesMissingButtonAppearanceFromExistingState() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildCheckboxWidgetWithoutAppearancePdf("Yes"));

        string output = Encoding.ASCII.GetString(flattened);
        string appearance = GetFlattenedAppearanceStreamText(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("1.25 w", appearance);
        Assert.Contains(" l S", appearance);
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
            Assert.Contains("(Path flatten) Tj", Encoding.ASCII.GetString(flattened));
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
            Assert.Contains("(Path stream flatten) Tj", Encoding.ASCII.GetString(flattenBytes));

            using var fillFlattenOutput = new MemoryStream();
            fillFlattenOutput.WriteByte(29);
            PdfFormFiller.FillAndFlattenFields(fillFlattenInputPath, fillFlattenOutput, new Dictionary<string, string> {
                ["Name"] = "Path stream single pass"
            });
            byte[] fillFlattenBytes = SliceAfterPrefix(fillFlattenOutput, 1);
            Assert.Equal(29, fillFlattenOutput.ToArray()[0]);
            Assert.False(PdfInspector.Inspect(fillFlattenBytes).HasForms);
            Assert.Contains("(Path stream single pass) Tj", Encoding.ASCII.GetString(fillFlattenBytes));
        } finally {
            if (File.Exists(fillInputPath)) File.Delete(fillInputPath);
            if (File.Exists(flattenInputPath)) File.Delete(flattenInputPath);
            if (File.Exists(fillFlattenInputPath)) File.Delete(fillFlattenInputPath);
        }
    }

    [Fact]
    public void FormPathOutputStreams_RejectNullAndReadOnlyOutputsBeforeReadingInputs() {
        var values = new Dictionary<string, string> {
            ["Name"] = "Value"
        };

        Assert.Throws<ArgumentNullException>(() => PdfFormFiller.FillFields("input.pdf", (Stream)null!, values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields("missing.pdf", new ReadOnlyStream(), values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(" ", new MemoryStream(), values));
        Assert.Throws<ArgumentNullException>(() => PdfFormFiller.FlattenFields("input.pdf", (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FlattenFields("missing.pdf", new ReadOnlyStream()));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FlattenFields(" ", new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfFormFiller.FillAndFlattenFields("input.pdf", (Stream)null!, values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillAndFlattenFields("missing.pdf", new ReadOnlyStream(), values));
        Assert.Throws<ArgumentException>(() => PdfFormFiller.FillAndFlattenFields(" ", new MemoryStream(), values));
    }

    [Fact]
    public void FillFields_RejectsUnknownFieldNames() {
        var ex = Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(BuildHierarchicalFormPdf(), new Dictionary<string, string> {
            ["Missing"] = "Value"
        }));

        Assert.Contains("PDF form field was not found: Missing", ex.Message);
    }

    [Fact]
    public void FillFields_RejectsSignedPdfs() {
        var ex = Assert.Throws<NotSupportedException>(() => PdfFormFiller.FillFields(BuildSignedFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Value"
        }));

        Assert.Equal("Signed PDF files are not supported for form filling by OfficeIMO.Pdf yet.", ex.Message);
    }

    [Fact]
    public void FlattenFields_RejectsSignedPdfs() {
        var ex = Assert.Throws<NotSupportedException>(() => PdfFormFiller.FlattenFields(BuildSignedFormPdf()));

        Assert.Equal("Signed PDF files are not supported for form flattening by OfficeIMO.Pdf yet.", ex.Message);
    }

    private static byte[] SliceAfterPrefix(MemoryStream stream, int prefixLength) {
        byte[] bytes = stream.ToArray();
        byte[] result = new byte[bytes.Length - prefixLength];
        Buffer.BlockCopy(bytes, prefixLength, result, 0, result.Length);
        return result;
    }

    private static byte[] BuildHierarchicalFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R 8 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /T (Person) /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /TU (Display name) /TM (ExportName) /V (OfficeIMO) /Ff 1 >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSignedFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R] /SigFlags 3 >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Name) /V (OfficeIMO) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCheckboxWidgetFormPdf() {
        string offAppearance = "% Unchecked appearance\n0.75 0.75 0.75 RG 0.5 0.5 15 15 re S";
        string checkedAppearance = "% Checked appearance\n0.75 0.75 0.75 RG 0.5 0.5 15 15 re S\n0 0 0 RG 3 8 m 7 3 l 13 13 l S";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 36 116] /F 4 /AS /Yes /AP << /N << /Off 9 0 R /Yes 10 0 R >> >> >>",
            "endobj",
            "9 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 16 16] /Length {Encoding.ASCII.GetByteCount(offAppearance)} >>",
            "stream",
            offAppearance,
            "endstream",
            "endobj",
            "10 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 16 16] /Length {Encoding.ASCII.GetByteCount(checkedAppearance)} >>",
            "stream",
            checkedAppearance,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCheckboxWidgetWithoutAppearancePdf(string stateName = "Off") {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            $"<< /FT /Btn /T (AcceptTerms) /V /{stateName} /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            $"<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 36 116] /F 4 /AS /{stateName} >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildChoiceWidgetFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Ch /T (Country) /V (DE) /Opt [[(PL) (Poland)] [(DE) (Germany)] [/US (United States)]] /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 200 122] /F 4 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMultiSelectChoiceWidgetFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 260 180] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Ch /T (Country) /V [(PL) /US] /Ff 2097152 /Opt [[(PL) (Poland)] [(DE) (Germany)] [/US (United States)]] /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 220 122] /F 4 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTextWidgetFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /V (OfficeIMO) /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 180 120] /F 4 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTextWidgetFormPdfWithReferencedContentArray() {
        string existing = "BT /F1 12 Tf 20 150 Td (Existing) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 10 0 R /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {existing.Length} >>",
            "stream",
            existing,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /V (OfficeIMO) /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 180 120] /F 4 >>",
            "endobj",
            "10 0 obj",
            "[4 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string GetFlattenedAppearanceStreamText(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.First(indirect =>
            indirect.Value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page").Value);
        PdfDictionary resources = Assert.IsType<PdfDictionary>(page.Items["Resources"]);
        PdfDictionary xObjects = Assert.IsType<PdfDictionary>(resources.Items["XObject"]);
        PdfReference reference = Assert.IsType<PdfReference>(xObjects.Items["OfficeIMOForm1"]);
        PdfStream stream = Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
        return Encoding.ASCII.GetString(stream.Data);
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}
