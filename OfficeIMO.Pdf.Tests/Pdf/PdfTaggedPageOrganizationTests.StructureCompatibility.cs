using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTaggedPageOrganizationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MarkedContentReferencesRetainOwnershipWithAnOptionalType(bool explicitType) {
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(CreateTaggedSource(), null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, catalog.Items["StructTreeRoot"])!;
            var document = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["K"])!;
            foreach (var child in ((PdfArray)document.Items["K"]).Items) {
                var element = (PdfDictionary)PdfObjectLookup.Resolve(objects, child)!;
                var reference = new PdfDictionary();
                if (explicitType) reference.Items["Type"] = new PdfName("MCR");
                reference.Items["Pg"] = element.Items["Pg"];
                reference.Items["MCID"] = new PdfNumber(0);
                element.Items["K"] = reference;
            }
            return security.InfoObjectNumber;
        });
        var input = PdfDocument.Load(source);
        foreach (var (output, pages) in new[] {
            (input.Pages.Reorder(2, 1), 2), (input.Pages.Extract(2), 1),
            (input.Pages.Delete(2), 1), (input.Pages.Duplicate(1), 3),
            (input.Pages.Append(input), 4)
        }) AssertStructure(PdfReadDocument.Open(output.ToBytes()), pages);
        Assert.Equal(source, input.ToBytes());
    }

    [Theory]
    [InlineData(true, "Document")]
    [InlineData(false, "Document")]
    [InlineData(true, "Part")]
    [InlineData(false, "Part")]
    [InlineData(true, null)]
    [InlineData(false, null)]
    public void ImportedRootLanguageRetainsCatalogInheritanceAndExplicitOverrides(bool sourceDocumentRoot, string? targetRole) {
        byte[] source = SetStructureLanguage(CreateTaggedSource(), "fr-FR", sourceDocumentRoot ? "Document" : null, true);
        byte[] target = targetRole == null
            ? SetStructureLanguage(PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Plain target")).ToBytes(), "en-US", null, false)
            : SetStructureLanguage(CreateTaggedSource(), "en-US", targetRole, false);
        var sourceDocument = PdfDocument.Load(source); var targetDocument = PdfDocument.Load(target);
        var output = PdfReadDocument.Open(targetDocument.Pages.Insert(1, sourceDocument).ToBytes());
        var catalog = Catalog(output);
        Assert.Equal("en-US", ((PdfStringObj)catalog.Items["Lang"]).Value);
        var root = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, catalog.Items["StructTreeRoot"])!;
        var ids = (PdfArray)((PdfDictionary)PdfObjectLookup.Resolve(output.Objects, root.Items["IDTree"])!).Items["Names"];
        foreach (var (pageIndex, expectedLanguage) in new[] { (0, "de-DE"), (1, "fr-FR") }) {
            var owner = ids.Items.OfType<PdfReference>()
                .Select(reference => (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, reference)!)
                .Single(element => element.Get<PdfReference>("Pg")?.ObjectNumber == output.Pages[pageIndex].ObjectNumber);
            Assert.Equal(expectedLanguage, EffectiveLanguage(owner));
        }
        Assert.Equal(source, sourceDocument.ToBytes()); Assert.Equal(target, targetDocument.ToBytes());

        string EffectiveLanguage(PdfDictionary element) {
            while (true) {
                if (element.Items.TryGetValue("Lang", out var language)) return ((PdfStringObj)PdfObjectLookup.Resolve(output.Objects, language)!).Value;
                if (!element.Items.TryGetValue("P", out var parent)) return ((PdfStringObj)catalog.Items["Lang"]).Value;
                element = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, parent)!;
            }
        }
    }

    private static byte[] SetStructureLanguage(byte[] source, string language, string? rootRole, bool overrideFirst) =>
        PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            catalog.Items["Lang"] = new PdfStringObj(language);
            if (catalog.Items.TryGetValue("StructTreeRoot", out var rootValue)) {
                var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, rootValue)!;
                var document = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["K"])!;
                var children = (PdfArray)document.Items["K"];
                if (overrideFirst) ((PdfDictionary)PdfObjectLookup.Resolve(objects, children.Items[0])!).Items["Lang"] = new PdfStringObj("de-DE");
                if (rootRole != null) document.Items["S"] = new PdfName(rootRole);
                else {
                    root.Items["K"] = children;
                    foreach (var child in children.Items) ((PdfDictionary)PdfObjectLookup.Resolve(objects, child)!).Items["P"] = rootValue;
                }
            }
            return security.InfoObjectNumber;
        });
}
