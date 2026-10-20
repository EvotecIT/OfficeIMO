using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTaggedMergeReadBudgetTests {
    [Fact]
    public void MergeAccountsForGeneratedClassObjectsAfterAdmittingBoundedSources() {
        byte[][] sources = { CreateReportWithIndependentClassArrays("Primary"), CreateReportWithIndependentClassArrays("Incoming") };
        byte[][] snapshots = sources.Select(bytes => (byte[])bytes.Clone()).ToArray();
        PdfDocument[] boundedDocuments = sources.Select(bytes => PdfDocument.Load(bytes, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxIndirectObjects = PdfReadDocument.Open(bytes).Objects.Count }
        })).ToArray();
        PdfMergeResult expected = PdfDocument.MergeResult(new PdfMergeOptions(), sources.Select(bytes => PdfDocument.Load(bytes)));

        PdfMergeResult merged = PdfDocument.MergeResult(new PdfMergeOptions(), boundedDocuments);

        byte[] output = merged.ToBytes();
        PdfReadDocument readback = PdfReadDocument.Open(output);
        Assert.Equal(4, merged.Report.OutputPageCount);
        Assert.Equal(4, readback.Pages.Count);
        Assert.Equal(80, readback.Objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Count(dictionary => dictionary.Items.ContainsKey("C")));
        Assert.Equal(expected.ToBytes(), output);
        for (int index = 0; index < sources.Length; index++) {
            Assert.Equal(snapshots[index], sources[index]);
            Assert.Equal(snapshots[index], boundedDocuments[index].ToBytes());
        }
    }

    private static byte[] CreateReportWithIndependentClassArrays(string label) {
        PdfDocument document = PdfDocument.Create().TaggedPdfCatalogMarkers().Language("en-US");
        for (int index = 0; index < 40; index++) document.Paragraph(paragraph => paragraph.Text(label + " ordinary paragraph " + index));
        return PdfDocumentObjectGraphRewriter.Rewrite(document.ToBytes(), null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, catalog.Items["StructTreeRoot"])!;
            var classes = new PdfDictionary();
            var attributes = new PdfDictionary();
            attributes.Items["O"] = new PdfName("Layout");
            classes.Items["Shared"] = attributes;
            root.Items["ClassMap"] = classes;
            var elements = objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
                .Where(dictionary => dictionary.Get<PdfName>("Type")?.Name == "StructElem" &&
                    dictionary.Get<PdfName>("S")?.Name != "Document").ToArray();
            int nextObjectNumber = objects.Keys.Max() + 1;
            foreach (var element in elements) {
                var classArray = new PdfArray();
                classArray.Items.Add(new PdfName("Shared"));
                classArray.Items.Add(new PdfNumber(1));
                objects.Add(nextObjectNumber, new PdfIndirectObject(nextObjectNumber, 0, classArray));
                element.Items["C"] = new PdfReference(nextObjectNumber++, 0);
            }
            return security.InfoObjectNumber;
        });
    }
}
