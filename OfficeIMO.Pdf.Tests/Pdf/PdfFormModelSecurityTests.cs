using System.Collections;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormModelSecurityTests {
    [Fact]
    public void PreflightCapabilityNumericContractsRemainStable() {
        Assert.Equal(0, (int)PdfPreflightCapability.ExtractText);
        Assert.Equal(1, (int)PdfPreflightCapability.ManipulatePages);
        Assert.Equal(2, (int)PdfPreflightCapability.FillSimpleFormFields);
        Assert.Equal(3, (int)PdfPreflightCapability.FlattenSimpleFormFields);
        Assert.Equal(4, (int)PdfPreflightCapability.FillAndFlattenSimpleFormFields);
        Assert.Equal(5, (int)PdfPreflightCapability.ExtractImages);
        Assert.Equal(6, (int)PdfPreflightCapability.ReadLogicalObjects);
        Assert.Equal(7, (int)PdfPreflightCapability.ExtractAttachments);
        Assert.Equal(8, (int)PdfPreflightCapability.AppendMetadataRevision);
        Assert.Equal(9, (int)PdfPreflightCapability.AppendFormFieldRevision);
        Assert.Equal(10, (int)PdfPreflightCapability.PrepareExternalSignatureRevision);
    }

    [Fact]
    public void FormFieldDerivedCollectionsRemainLinearAndPreserveOrder() {
        var widgets = new List<PdfFormWidget>();
        for (int pageNumber = 1; pageNumber <= 4096; pageNumber++) {
            widgets.Add(new PdfFormWidget(pageNumber, "Choice", pageNumber, 0, 0, 10, 10, null, null));
            widgets.Add(new PdfFormWidget(pageNumber + 4096, "Choice", pageNumber, 0, 0, 10, 10, null, null));
        }

        var values = new List<string>();
        var options = new List<PdfFormFieldOption>();
        for (int index = 0; index < 4096; index++) {
            string value = "Value" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            values.Add(value);
            options.Add(new PdfFormFieldOption(value, "Display " + value));
        }

        var field = new PdfFormField(
            objectNumber: 1,
            name: "Choice",
            partialName: "Choice",
            fieldType: "Ch",
            value: values[0],
            alternateName: null,
            mappingName: null,
            flags: null,
            values: values,
            options: options,
            widgets: widgets);

        Assert.Equal(Enumerable.Range(1, 4096), field.PageNumbers);
        Assert.Equal(options, field.SelectedOptions);
    }

    [Fact]
    public void LogicalWidgetIndexBuildsEachWidgetOnceByPage() {
        var fields = new List<PdfFormField>();
        for (int fieldIndex = 0; fieldIndex < 1024; fieldIndex++) {
            int pageNumber = fieldIndex % 16 + 1;
            var widget = new PdfFormWidget(fieldIndex, "Field" + fieldIndex, pageNumber, 0, 0, 10, 10, null, null);
            fields.Add(new PdfFormField(
                objectNumber: fieldIndex,
                name: "Field" + fieldIndex,
                partialName: "Field" + fieldIndex,
                fieldType: "Tx",
                value: null,
                alternateName: null,
                mappingName: null,
                flags: null,
                widgets: new[] { widget }));
        }

        IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalFormWidget>> index =
            PdfLogicalDocument.IndexFormWidgetsByPageNumber(fields);

        Assert.Equal(16, index.Count);
        Assert.Equal(1024, index.Values.Sum(widgets => widgets.Count));
        Assert.All(index, pair => Assert.All(pair.Value, widget => Assert.Equal(pair.Key, widget.PageNumber)));
    }

    [Fact]
    public void StructuredLeaderRowsDeduplicateWithStableInsertionOrder() {
        var page = new StructuredPage();
        for (int index = 0; index < 4096; index++) {
            Assert.True(page.TryAddLeaderRow("Label " + index, "Value " + index));
        }

        Assert.False(page.TryAddLeaderRow("Label 2048", "Value 2048"));
        Assert.Equal(4096, page.LeaderRows.Count);
        Assert.Equal(new[] { "Label 0", "Value 0" }, page.LeaderRows[0]);
        Assert.Equal(new[] { "Label 4095", "Value 4095" }, page.LeaderRows[4095]);
    }

    [Fact]
    public void XrefStreamRejectsOversizedFieldWidthsWithoutEnumeratingBytes() {
        var widths = new PdfArray();
        widths.Items.Add(new PdfNumber(int.MaxValue));
        widths.Items.Add(new PdfNumber(int.MaxValue));
        widths.Items.Add(new PdfNumber(3));
        var dictionary = new PdfDictionary();
        dictionary.Items["W"] = widths;
        dictionary.Items["Size"] = new PdfNumber(1);
        MethodInfo method = typeof(PdfSyntax).GetMethod(
            "ReadXrefStreamEntries",
            BindingFlags.NonPublic | BindingFlags.Static)!;

        var entries = (IEnumerable)method.Invoke(null, new object[] { dictionary, new byte[] { 0, 0, 0 } })!;

        Assert.Empty(entries.Cast<object>());
    }

    [Fact]
    public void LiteralStringUsesUnicodeEncodingBeforeLowBytesCanBecomePdfSyntax() {
        string encoded = PdfSyntaxEscaper.LiteralString("\u0129 /Author \u0128Injected");

        Assert.StartsWith("<", encoded, StringComparison.Ordinal);
        Assert.EndsWith(">", encoded, StringComparison.Ordinal);
        Assert.DoesNotContain(") /Author (", encoded, StringComparison.Ordinal);
    }
}
