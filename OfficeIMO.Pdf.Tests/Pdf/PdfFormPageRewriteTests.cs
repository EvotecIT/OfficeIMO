using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig.Tokens;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.AcroForms;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormPageRewriteTests {
    [Theory]
    [InlineData("insert")]
    [InlineData("extract")]
    [InlineData("reorder")]
    [InlineData("move")]
    public void ReportLabFieldsRemainLiveAfterPageTransfer(string operation) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "PdfFormOcr", "reportlab-scanned-form.pdf"));
        PdfDocument source = PdfDocument.Load(bytes);
        PdfDocument result = operation switch {
            "insert" => PdfDocument.Create(compose => { compose.Page(page => page.Size(300, 400)); compose.Page(page => page.Size(300, 400)); }).Pages.Insert(2, source),
            "extract" => source.Pages.Extract(1),
            "reorder" => source.Pages.Reorder(2, 1),
            _ => source.Pages.Move(3, 1)
        };
        using PdfPigDocument independent = PdfPigDocument.Open(result.ToBytes());
        Assert.True(independent.TryGetForm(out var form));
        var fields = form.GetFields().ToArray();
        Assert.Equal(operation == "extract" ? 4 : 7, fields.Length);
        if (operation != "extract") Assert.Equal("KEEP", result.Inspect().FormFieldsByName["ReadOnly"].Value);
        Assert.Equal(operation == "insert" ? 4 : operation == "extract" ? 1 : 2, independent.NumberOfPages);
    }

    [Theory]
    [InlineData("insert")]
    [InlineData("append")]
    [InlineData("prepend")]
    [InlineData("merge")]
    public void SequentialImportsKeepDistinctFieldsAndSharedWidgets(string operation) {
        PdfDocument source = PdfDocument.Load(CreateSharedWidgetForm());
        PdfDocument result = source;
        for (int index = 0; index < 2; index++) result = operation switch {
            "insert" => result.Pages.Insert(index + 2, source),
            "append" => result.Pages.Append(source),
            "prepend" => result.Pages.Prepend(source),
            _ => PdfDocument.Merge(new PdfMergeOptions { Policy = new PdfMergePolicy { Forms = PdfMergeStructureMode.Combine } }, result, source)
        };
        var info = result.Inspect();
        Assert.Equal(new[] { "Shared", "Shared.source2", "Shared.source2.2" }, info.FormFieldNames.OrderBy(name => name).ToArray());
        Assert.All(info.FormFields, field => {
            Assert.Equal("Original", field.Value);
            Assert.Equal(2, field.Widgets.Count);
            Assert.All(field.Widgets, widget => Assert.True(widget.PageNumber.HasValue));
        });
        using PdfPigDocument independent = PdfPigDocument.Open(result.ToBytes());
        Assert.True(independent.TryGetForm(out var form));
        Assert.Equal(3, form.Fields.Count);
        Assert.True(form.NeedAppearances);
        var calculations = Assert.IsType<ArrayToken>(form.Dictionary.Data[NameToken.Create("CO")]);
        Assert.Equal(Assert.IsType<ArrayToken>(form.Dictionary.Data[NameToken.Create("Fields")]).Data
            .Select(token => Assert.IsType<IndirectReferenceToken>(token).Data).OrderBy(reference => reference.ObjectNumber), calculations.Data
            .Select(token => Assert.IsType<IndirectReferenceToken>(token).Data)
            .OrderBy(reference => reference.ObjectNumber));
        var defaultResources = Assert.IsType<DictionaryToken>(form.Dictionary.Data[NameToken.Create("DR")]);
        var fonts = Assert.IsType<DictionaryToken>(defaultResources.Data[NameToken.Create("Font")]);
        foreach (var field in form.Fields) {
            var value = field.Dictionary.Data[NameToken.Create("DA")];
            string appearance = value is StringToken literal ? literal.Data : Assert.IsType<HexToken>(value).Data;
            var font = Regex.Match(appearance, @"/([^\s]+)\s+[0-9.]+\s+Tf");
            Assert.True(font.Success, "Default appearance must contain a PDF font-name operand: " + appearance);
            Assert.Contains(NameToken.Create(font.Groups[1].Value), fonts.Data.Keys);
        }
        if (operation == "insert") {
            Assert.Equal(new int?[] { 1, 6 }, info.FormFieldsByName["Shared"].Widgets.Select(widget => widget.PageNumber));
            Assert.Equal(new int?[] { 2, 5 }, info.FormFieldsByName["Shared.source2"].Widgets.Select(widget => widget.PageNumber));
            Assert.Equal(new int?[] { 3, 4 }, info.FormFieldsByName["Shared.source2.2"].Widgets.Select(widget => widget.PageNumber));
        }
        PdfDocument filled = result.Forms.Fill(new Dictionary<string, string> { ["Shared.source2"] = "Changed" });
        Assert.Equal("Original", filled.Inspect().FormFieldsByName["Shared"].Value);
        Assert.Equal("Changed", filled.Inspect().FormFieldsByName["Shared.source2"].Value);
        PdfDocument partial = result.Pages.Extract(operation == "prepend" ? 1 : 2);
        Assert.Single(partial.Inspect().FormFields);
        Assert.Single(Assert.Single(partial.Inspect().FormFields).Widgets);
        if (operation == "insert") {
            using var repeated = PdfPigDocument.Open(partial.Pages.Extract(1, 1).ToBytes());
            Assert.True(repeated.TryGetForm(out var repeatedForm));
            Assert.Equal(2, repeatedForm.Fields.Count);
            var order = Assert.IsType<ArrayToken>(repeatedForm.Dictionary.Data[NameToken.Create("CO")]);
            Assert.Equal(2, order.Data.Distinct().Count());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportRefusesFieldOwnershipThatCannotBePreserved(bool taggedTarget) {
        var source = PdfDocument.Load(File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "PdfFormOcr", "reportlab-scanned-form.pdf")));
        PdfDocument target = taggedTarget ? PdfDocument.Create().TaggedPdfCatalogMarkers().H1("Tagged") : source;
        byte[] before = target.ToBytes();
        var error = Assert.Throws<NotSupportedException>(() => target.Pages.Insert(1, source));
        Assert.Contains(taggedTarget ? "structure ownership" : "field-name references", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, target.ToBytes());
    }

    private static byte[] CreateSharedWidgetForm() => BuildRawPdf(
        "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
        "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << >> /Annots [7 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << >> /Annots [8 0 R] >>",
        "<< /Fields [6 0 R] /CO [6 0 R] /NeedAppearances true /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 9 0 R >> >> >>",
        "<< /FT /Tx /T (Shared) /V (Original) /Kids [7 0 R 8 0 R] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 6 0 R /P 3 0 R /Rect [20 20 180 50] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 6 0 R /P 4 0 R /Rect [20 20 180 50] >>",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

    private static byte[] BuildRawPdf(params string[] objectBodies) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int> { 0 };
        for (int objectIndex = 0; objectIndex < objectBodies.Length; objectIndex++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(objectIndex + 1).Append(" 0 obj\n")
                .Append(objectBodies[objectIndex]).Append("\nendobj\n");
        }

        int xref = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(offsets.Count).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets.Skip(1)) builder.Append(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(offsets.Count)
            .Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
