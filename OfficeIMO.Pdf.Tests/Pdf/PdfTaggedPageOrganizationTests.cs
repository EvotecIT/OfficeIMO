using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTaggedPageOrganizationTests {
    [Fact]
    public void CopiesAndImportsRetainOpaqueStructureIdsAndRemapStructuredDestinations() {
        byte[] source = CreateTaggedSource();
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, catalog.Items["StructTreeRoot"])!;
            var document = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["K"])!;
            var element = (PdfDictionary)PdfObjectLookup.Resolve(objects, ((PdfArray)document.Items["K"]).Items[0])!;
            var id = new PdfStringObj(new byte[] { 0, 255, 1 }); element.Items["ID"] = id;
            var ids = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["IDTree"])!;
            ((PdfArray)ids.Items["Names"]).Items[0] = id;
            var page = (PdfDictionary)PdfObjectLookup.Resolve(objects, element.Items["Pg"])!;
            var action = new PdfDictionary(); action.Items["S"] = new PdfName("GoTo"); action.Items["SD"] = id;
            var destination = new PdfArray(); destination.Items.Add(element.Items["Pg"]); destination.Items.Add(new PdfName("Fit")); action.Items["D"] = destination;
            ((PdfDictionary)page.Items["AA"]).Items["C"] = action;
            return security.InfoObjectNumber;
        });
        var input = PdfDocument.Load(source);
        foreach (var (output, expectedDestinations) in new[] {
            (input.Pages.Duplicate(1), 2), (input.Pages.Insert(1, input), 2),
            (input.Pages.Duplicate(1).Pages.Duplicate(1), 3),
            (input.Pages.Append(input).Pages.Append(input), 3)
        }) {
            var read = PdfReadDocument.Open(output.ToBytes()); var objects = read.Objects;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, Catalog(read).Items["StructTreeRoot"])!;
            var names = (PdfArray)((PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["IDTree"])!).Items["Names"];
            var owners = new Dictionary<string, PdfDictionary>(StringComparer.Ordinal);
            for (int index = 0; index < names.Items.Count; index += 2)
                owners.Add(Convert.ToBase64String(((PdfStringObj)names.Items[index]).RawBytes), (PdfDictionary)PdfObjectLookup.Resolve(objects, names.Items[index + 1])!);
            int destinations = 0;
            foreach (var page in read.Pages) if (page.PageDictionary.Items.TryGetValue("AA", out var aaValue) &&
                PdfObjectLookup.Resolve(objects, aaValue) is PdfDictionary aa && aa.Items.TryGetValue("C", out var actionValue)) {
                var action = (PdfDictionary)PdfObjectLookup.Resolve(objects, actionValue)!;
                var id = (PdfStringObj)action.Items["SD"]; var owner = owners[Convert.ToBase64String(id.RawBytes)];
                Assert.Equal(new byte[] { 0, 255, 1 }, id.RawBytes.Skip(id.RawBytes.Length - 3));
                Assert.Equal(page.ObjectNumber, ((PdfReference)owner.Items["Pg"]).ObjectNumber); destinations++;
            }
            Assert.Equal(expectedDestinations, destinations);
        }
        Assert.Equal(source, input.ToBytes());
    }

    [Fact]
    public void CopyingDirectDocumentMarkedContentCreatesASeparatePageOwner() {
        byte[] source = PdfDocument.Load(CreateTaggedSource()).Pages.Extract(1).ToBytes();
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, catalog.Items["StructTreeRoot"])!;
            var documentReference = (PdfReference)root.Items["K"];
            var document = (PdfDictionary)objects[documentReference.ObjectNumber].Value;
            var leafReference = (PdfReference)((PdfArray)document.Items["K"]).Items[0];
            var leaf = (PdfDictionary)objects[leafReference.ObjectNumber].Value;
            document.Items["Pg"] = leaf.Items["Pg"]; document.Items["K"] = leaf.Items["K"];
            var parents = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["ParentTree"])!;
            ((PdfArray)((PdfArray)parents.Items["Nums"]).Items[1]).Items[0] = documentReference;
            root.Items.Remove("IDTree");
            return security.InfoObjectNumber;
        });
        var output = PdfReadDocument.Open(PdfDocument.Load(source).Pages.Duplicate(1).ToBytes());
        var rootAfter = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, Catalog(output).Items["StructTreeRoot"])!;
        var nums = (PdfArray)((PdfDictionary)rootAfter.Items["ParentTree"]).Items["Nums"];
        Assert.Equal(4, nums.Items.Count);
        for (int index = 0; index < nums.Items.Count; index += 2) {
            var owner = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, ((PdfArray)nums.Items[index + 1]).Items[0])!;
            var page = output.Pages.Single(item => item.PageDictionary.Get<PdfNumber>("StructParents")!.Value == ((PdfNumber)nums.Items[index]).Value);
            Assert.Equal(page.ObjectNumber, ((PdfReference)owner.Items["Pg"]).ObjectNumber);
        }
    }

    [Fact]
    public void CopyingAFormPageClonesItsFieldOwnershipAndDeletionPrunesOnlyTheRemovedWidgets() {
        var source = PdfDocument.Create().TextField("First", 120, 24, "First value").PageBreak().TextField("Second", 120, 24, "Second value");
        PdfDocument copied = source.Pages.Duplicate(1); var fields = copied.Inspect().FormFields;
        Assert.Equal(3, fields.Count); Assert.Equal(3, fields.Select(field => field.Name).Distinct().Count());
        Assert.Equal(new[] { 1, 2, 3 }, fields.SelectMany(field => field.Widgets).Select(widget => widget.PageNumber!.Value).OrderBy(page => page));
        PdfDocument deleted = copied.Pages.Delete(1);
        Assert.Equal(2, deleted.Inspect().FormFields.Count);
        Assert.DoesNotContain(deleted.Inspect().FormFields, field => field.Name == "First");
        Assert.Contains(deleted.Inspect().FormFields, field => field.Name != null && field.Name.StartsWith("First_copy", StringComparison.Ordinal));
        Assert.Equal(2, source.Inspect().FormFields.Count);
    }

    [Theory]
    [InlineData(PdfStandardPermissions.None, false)]
    [InlineData(PdfStandardPermissions.AssembleDocument, true)]
    [InlineData(PdfStandardPermissions.CopyContents, false)]
    public void TaggedOrganizationRetainsEncryptionAndChecksAssemblySeparatelyFromCopy(PdfStandardPermissions permissions, bool allowed) {
        byte[] bytes = PdfDocumentObjectGraphRewriter.Rewrite(CreateTaggedSource(), null, new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner", AllowedPermissions = permissions
        });
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "open" });
        var result = document.Pages.ReorderResult("2,1");
        Assert.True(allowed == result.Succeeded, result.Exception?.ToString() ?? result.MutationPlan?.ToString());
        if (allowed) {
            var output = result.RequireValue(); var read = PdfReadDocument.Open(output.ToBytes(), new PdfLoadOptions { Password = "open" });
            Assert.True(read.Security.HasEncryption); AssertStructure(read, 2);
            var rotate = document.Pages.RotateResult(180, PdfPageSelector.Parse("last"));
            Assert.True(rotate.Succeeded, rotate.Exception?.ToString() ?? string.Join(" ", rotate.Diagnostics));
            var crop = document.Pages.SetPageBoxResult(PdfPageBoundaryBox.CropBox, 10, 10, 280, 480, PdfPageSelection.From(1));
            Assert.False(crop.Succeeded); // Assembly permits rotation; changing visible page boundaries requires modification.
        }
        Assert.False(document.PlanMutation(PdfMutationOperation.ImportPages).CanExecute);
        Assert.Equal(bytes, document.ToBytes());
    }

    [Theory]
    [InlineData(1)] [InlineData(2)] [InlineData(3)]
    public void CertificationStillBlocksCatalogPreservingOrganization(int permission) {
        byte[] source = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permission);
        foreach (var operation in new[] { PdfMutationOperation.OrganizePages, PdfMutationOperation.ImportPages })
            Assert.False(PdfMutationPlanner.Plan(source, operation).CanExecute);
    }

    [Fact]
    public void ImportCombinesStructureParentsAndNamespacesIdsWithoutChangingEitherInput() {
        byte[] bytes = CreateTaggedSource(); PdfDocument target = PdfDocument.Load(bytes), source = PdfDocument.Load(bytes);
        var result = target.Pages.InsertResult(2, source);
        var output = result.RequireValue(); var read = PdfReadDocument.Open(output.ToBytes());
        Assert.Equal(PdfMutationOperation.ImportPages, result.MutationPlan!.Operation);
        Assert.Contains(PdfMutationPermissionCheck.CopyContents, result.MutationPlan.PermissionChecks);
        Assert.Equal(4, read.Pages.Count); AssertStructure(read, 4);
        Assert.Equal(bytes, target.ToBytes()); Assert.Equal(bytes, source.ToBytes());
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output.ToBytes());
        Assert.Contains("First", independent.GetPage(1).Text); Assert.Contains("First", independent.GetPage(2).Text);
        Assert.Contains("Second", independent.GetPage(3).Text); Assert.Contains("Second", independent.GetPage(4).Text);
        Assert.Equal(2, output.Inspect().Outlines.Count);
    }

    [Theory]
    [InlineData(true)] [InlineData(false)]
    public void MixedTaggedAndPlainInsertionRetainsTheTaggedStructure(bool taggedTarget) {
        var tagged = PdfDocument.Load(CreateTaggedSource());
        var plain = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 500).Content(content => content.Text("Plain inserted page"))));
        PdfDocument output = taggedTarget ? tagged.Pages.Insert(2, plain) : plain.Pages.Insert(1, tagged);
        Assert.Equal(3, output.Inspect().PageCount); Assert.True(output.Inspect().HasTaggedContent);
        AssertStructure(PdfReadDocument.Open(output.ToBytes()), 2);
        Assert.Contains("Plain inserted page", PdfTextExtractor.ExtractAllText(output.ToBytes()));
    }

    [Fact]
    public void DuplicateCreatesIndependentPageStructureKeysAndIdsForEveryCopy() {
        byte[] source = CreateTaggedSource(); PdfDocument document = PdfDocument.Load(source);
        var result = document.Pages.DuplicateResult("1,1");
        PdfDocument output = result.RequireValue(); var read = PdfReadDocument.Open(output.ToBytes());
        AssertStructure(read, 4);
        Assert.Equal(4, read.Pages.Select(page => page.PageDictionary.Get<PdfNumber>("StructParents")!.Value).Distinct().Count());
        Assert.Equal(source, document.ToBytes());
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output.ToBytes());
        Assert.Equal(4, independent.NumberOfPages);
        Assert.All(Enumerable.Range(1, 3), page => Assert.Contains("First", independent.GetPage(page).Text));
        Assert.Contains("Second", independent.GetPage(4).Text);
    }

    [Fact]
    public void ReorderRetainsCatalogStructureOwnershipInheritanceAndObjectMapping() {
        byte[] source = CreateTaggedSource(); var before = PdfReadDocument.Open(source);
        PdfDocument document = PdfDocument.Load(source);
        var result = document.Pages.ReorderWithMapping(2, 1);
        var after = PdfReadDocument.Open(result.Bytes);
        Assert.Equal(source, document.ToBytes());
        Assert.Equal(after.Pages[0].ObjectNumber, result.ObjectNumberMap[before.Pages[1].ObjectNumber]);
        var catalog = Catalog(after);
        Assert.Equal("en-US", ((PdfStringObj)catalog.Items["Lang"]).Value);
        Assert.True(((PdfBoolean)((PdfDictionary)catalog.Items["MarkInfo"]).Items["Marked"]).Value);
        Assert.Equal(90, after.Pages[0].GetRotationDegrees());
        Assert.Equal(90, after.Pages[1].GetRotationDegrees());
        AssertStructure(after, 2);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(result.Bytes);
        Assert.Contains("Second", independent.GetPage(1).Text);
        Assert.Contains("First", independent.GetPage(2).Text);
    }

    [Fact]
    public void DeletePrunesOnlyRemovedTagsParentEntriesIdsBookmarksAndIncomingActions() {
        PdfDocument source = PdfDocument.Load(CreateTaggedSource());
        var result = source.Pages.DeleteResult(PdfPageSelection.From(2));
        var output = result.RequireValue(); var read = PdfReadDocument.Open(output.ToBytes());
        Assert.Equal(PdfMutationOperation.OrganizePages, result.MutationPlan!.Operation);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.Single(read.Pages); AssertStructure(read, 1);
        Assert.Equal("First bookmark", output.Inspect().Outlines.Single().Title);
        var page = (PdfDictionary)read.Objects[read.Pages[0].ObjectNumber].Value;
        var actions = (PdfDictionary)page.Items["AA"];
        Assert.Empty(actions.Items);
        Assert.False(Catalog(read).Items.ContainsKey("OpenAction"));
        Assert.Contains("First", PdfTextExtractor.ExtractAllText(output.ToBytes()));
    }

    [Fact]
    public void CompoundExtractionCreatesIndependentPrunedStructureGraphs() {
        byte[] source = CreateTaggedSource(); PdfDocument document = PdfDocument.Load(source);
        var parts = document.Pages.Split(PdfPageSelection.From(1), PdfPageSelection.From(2));
        Assert.Equal(source, document.ToBytes());
        Assert.Equal(2, parts.Count);
        foreach (var part in parts) AssertStructure(PdfReadDocument.Open(part.ToBytes()), 1);
        Assert.Contains("First", PdfTextExtractor.ExtractAllText(parts[0].ToBytes()));
        Assert.Contains("Second", PdfTextExtractor.ExtractAllText(parts[1].ToBytes()));
    }

    [Fact]
    public void IndependentTaggedProducerRetainsTagsOnReorderAndExtract() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-pdfa1b-text.pdf");
        PdfDocument source = PdfDocument.Load(path);
        Assert.True(source.Inspect().HasTaggedContent);
        foreach (PdfDocument output in new[] { source.Pages.Reorder(1), source.Pages.Extract(1), source.Pages.Duplicate(1) }) {
            Assert.True(output.Inspect().HasTaggedContent);
            using var independent = UglyToad.PdfPig.PdfDocument.Open(output.ToBytes());
            Assert.Equal(output.Inspect().PageCount, independent.NumberOfPages);
            Assert.NotEmpty(independent.GetPage(1).Letters);
        }
    }

    [Fact]
    public void UnrecognizedSurvivingPageReferenceFailsWithoutChangingSource() {
        byte[] source = CreateTaggedSource(); var initial = PdfReadDocument.Open(source);
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            ((PdfDictionary)objects[security.RootObjectNumber!.Value].Value).Items["CustomPageOwner"] = new PdfReference(initial.Pages[1].ObjectNumber, 0);
            return security.InfoObjectNumber;
        });
        PdfDocument document = PdfDocument.Load(source);
        var result = document.Pages.DeleteResult(PdfPageSelection.From(2));
        Assert.False(result.Succeeded); Assert.Equal(source, document.ToBytes());
        Assert.Contains("surviving reference", result.Exception!.Message);
    }

    private static void AssertStructure(PdfReadDocument read, int pages) {
        var objects = read.Objects; var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, Catalog(read).Items["StructTreeRoot"])!;
        var leaves = new List<PdfDictionary>(); Visit(root.Items["K"], (PdfReference)Catalog(read).Items["StructTreeRoot"]);
        Assert.Equal(pages, leaves.Count);
        var parents = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["ParentTree"])!;
        var nums = (PdfArray)parents.Items["Nums"]; Assert.Equal(pages * 2, nums.Items.Count);
        var ids = (PdfDictionary)PdfObjectLookup.Resolve(objects, root.Items["IDTree"])!;
        Assert.Equal(pages * 2, ((PdfArray)ids.Items["Names"]).Items.Count);
        var pageIds = new HashSet<int>(read.Pages.Select(page => page.ObjectNumber));
        foreach (var element in leaves) {
            Assert.Contains(((PdfReference)element.Items["Pg"]).ObjectNumber, pageIds);
            var content = PdfObjectLookup.Resolve(objects, element.Items["K"]);
            if (content is PdfNumber number) Assert.Equal(0, number.Value);
            else {
                var reference = Assert.IsType<PdfDictionary>(content);
                Assert.Equal(0, reference.Get<PdfNumber>("MCID")!.Value);
                Assert.Equal(((PdfReference)element.Items["Pg"]).ObjectNumber, reference.Get<PdfReference>("Pg")!.ObjectNumber);
            }
        }
        for (int index = 0; index < nums.Items.Count; index += 2) {
            int key = (int)((PdfNumber)nums.Items[index]).Value;
            var page = read.Pages.Single(item => ((PdfDictionary)objects[item.ObjectNumber].Value).Get<PdfNumber>("StructParents")?.Value == key);
            var owner = (PdfDictionary)PdfObjectLookup.Resolve(objects, ((PdfArray)nums.Items[index + 1]).Items[0])!;
            Assert.Equal(page.ObjectNumber, ((PdfReference)owner.Items["Pg"]).ObjectNumber);
        }
        void Visit(PdfObject item, PdfReference parent) {
            if (PdfObjectLookup.Resolve(objects, item) is PdfArray array) { foreach (var child in array.Items) Visit(child, parent); return; }
            var reference = Assert.IsType<PdfReference>(item); var element = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, reference));
            Assert.Equal(parent.ObjectNumber, ((PdfReference)element.Items["P"]).ObjectNumber);
            var content = PdfObjectLookup.Resolve(objects, element.Items["K"]);
            if (content is PdfNumber || content is PdfDictionary dictionary && dictionary.Items.ContainsKey("MCID")) leaves.Add(element);
            else Visit(element.Items["K"], reference);
        }
    }

    internal static byte[] CreateTaggedSource() {
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Size(300, 500).Content(content => content.Text("First tagged page")));
            compose.Page(page => page.Size(300, 500).Content(content => content.Text("Second tagged page")));
        }).ToBytes();
        return PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var parsed = PdfReadDocument.Open(source); var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            int next = objects.Keys.Max() + 1;
            PdfReference Add(PdfObject value) { int number = next++; objects.Add(number, new PdfIndirectObject(number, 0, value)); return new PdfReference(number, 0); }
            var root = new PdfDictionary(); root.Items["Type"] = new PdfName("StructTreeRoot"); var rootRef = Add(root);
            var document = new PdfDictionary(); document.Items["Type"] = new PdfName("StructElem"); document.Items["S"] = new PdfName("Document");
            document.Items["P"] = rootRef; var documentRef = Add(document); root.Items["K"] = documentRef;
            var children = new PdfArray(); document.Items["K"] = children;
            var nums = new PdfArray(); var ids = new PdfArray();
            var outlineRoot = new PdfDictionary(); outlineRoot.Items["Type"] = new PdfName("Outlines"); var outlineRootRef = Add(outlineRoot);
            PdfDictionary? previous = null; PdfReference? previousRef = null;
            for (int index = 0; index < 2; index++) {
                int pageNumber = parsed.Pages[index].ObjectNumber; var page = (PdfDictionary)objects[pageNumber].Value;
                page.Items["StructParents"] = new PdfNumber(index);
                var element = new PdfDictionary(); element.Items["Type"] = new PdfName("StructElem"); element.Items["S"] = new PdfName("P");
                element.Items["P"] = documentRef; element.Items["Pg"] = new PdfReference(pageNumber, 0); element.Items["K"] = new PdfNumber(0);
                element.Items["ID"] = new PdfStringObj("p" + index); var elementRef = Add(element); children.Items.Add(elementRef);
                var parents = new PdfArray(); parents.Items.Add(elementRef); nums.Items.Add(new PdfNumber(index)); nums.Items.Add(parents);
                ids.Items.Add(element.Items["ID"]); ids.Items.Add(elementRef);
                var content = new PdfArray(); content.Items.Add(Add(new PdfStream(new PdfDictionary(), System.Text.Encoding.ASCII.GetBytes("/P << /MCID 0 >> BDC\n"))));
                if (PdfObjectLookup.Resolve(objects, page.Items["Contents"]) is PdfArray streams) content.Items.AddRange(streams.Items);
                else content.Items.Add(page.Items["Contents"]);
                content.Items.Add(Add(new PdfStream(new PdfDictionary(), System.Text.Encoding.ASCII.GetBytes("\nEMC\n")))); page.Items["Contents"] = content;
                var destination = new PdfArray(); destination.Items.Add(new PdfReference(pageNumber, 0)); destination.Items.Add(new PdfName("Fit"));
                var outline = new PdfDictionary(); outline.Items["Title"] = new PdfStringObj(index == 0 ? "First bookmark" : "Second bookmark");
                outline.Items["Dest"] = destination; outline.Items["Parent"] = outlineRootRef; var outlineRef = Add(outline);
                if (previous != null) { previous.Items["Next"] = outlineRef; outline.Items["Prev"] = previousRef!; }
                else outlineRoot.Items["First"] = outlineRef;
                outlineRoot.Items["Last"] = outlineRef; previous = outline; previousRef = outlineRef;
                if (index == 1) {
                    var action = new PdfDictionary(); action.Items["S"] = new PdfName("GoTo"); action.Items["D"] = destination;
                    var aa = new PdfDictionary(); aa.Items["O"] = Add(action); ((PdfDictionary)objects[parsed.Pages[0].ObjectNumber].Value).Items["AA"] = aa;
                    catalog.Items["OpenAction"] = destination;
                }
            }
            var parentTree = new PdfDictionary(); parentTree.Items["Nums"] = nums; root.Items["ParentTree"] = Add(parentTree); root.Items["ParentTreeNextKey"] = new PdfNumber(2);
            var idTree = new PdfDictionary(); idTree.Items["Names"] = ids; root.Items["IDTree"] = Add(idTree);
            outlineRoot.Items["Count"] = new PdfNumber(2); catalog.Items["Outlines"] = outlineRootRef;
            catalog.Items["StructTreeRoot"] = rootRef; catalog.Items["Lang"] = new PdfStringObj("en-US");
            var mark = new PdfDictionary(); mark.Items["Marked"] = new PdfBoolean(true); catalog.Items["MarkInfo"] = mark;
            var parent = (PdfReference)((PdfDictionary)objects[parsed.Pages[0].ObjectNumber].Value).Items["Parent"];
            ((PdfDictionary)objects[parent.ObjectNumber].Value).Items["Rotate"] = new PdfNumber(90);
            return security.InfoObjectNumber;
        });
    }
    private static PdfDictionary Catalog(PdfReadDocument read) => (PdfDictionary)read.Objects[read.Security.RootObjectNumber!.Value].Value;
}
