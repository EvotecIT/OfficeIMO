using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTaggedMergeTests {
    [Fact]
    public void MergeRetainsIndependentProducerTagsBesideAnOwnedReport() {
        string path = Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-pdfa1b-text.pdf");
        byte[] external = File.ReadAllBytes(path);
        byte[] owned = Report("Service", "en-US");
        byte[] output = PdfDocument.MergeToBytes(owned, external);
        using var sourceReader = UglyToad.PdfPig.PdfDocument.Open(external);
        using var outputReader = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Equal(sourceReader.NumberOfPages + 1, outputReader.NumberOfPages);
        for (int page = 1; page <= sourceReader.NumberOfPages; page++)
            Assert.Equal(sourceReader.GetPage(page).Text, outputReader.GetPage(page + 1).Text);
        AssertParentTreeMatchesPages(PdfReadDocument.Open(output), outputReader.NumberOfPages);
        Assert.Equal(external, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData(PdfStandardPermissions.None, false)]
    [InlineData(PdfStandardPermissions.AssembleDocument, false)]
    [InlineData(PdfStandardPermissions.CopyContents, false)]
    [InlineData(PdfStandardPermissions.CopyContents | PdfStandardPermissions.AssembleDocument, true)]
    public void MergeChecksCopyAndAssemblyPermissionsForTaggedIncomingPages(PdfStandardPermissions permissions, bool allowed) {
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(Report("Private", "en-US"), null,
            new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner", AllowedPermissions = permissions });
        PdfDocument encrypted = PdfDocument.Load(source, new PdfLoadOptions { Password = "open" });
        PdfDocument primary = PdfDocument.Load(Report("Service", "en-US"));
        if (allowed) {
            PdfDocument output = PdfDocument.Merge(primary, encrypted);
            Assert.False(output.Inspect().Security.HasEncryption);
            AssertParentTreeMatchesPages(PdfReadDocument.Open(output.ToBytes()), 2);
        } else {
            var error = Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Merge(primary, encrypted));
            Assert.Contains(PdfMutationPermissionCheck.CopyContents, error.Plan.PermissionChecks);
            Assert.Contains(PdfMutationPermissionCheck.AssembleDocument, error.Plan.PermissionChecks);
        }
        Assert.Equal(source, encrypted.ToBytes());
    }

    [Fact]
    public void MergeRejectsUntaggedPagesWhenPrimaryDeclaresPdfUa() {
        byte[] primary = PdfDocument.Create(new PdfCreateOptions().SetPdfUaIdentification())
            .TaggedPdfCatalogMarkers().H1("Service review").ToBytes();
        byte[] plain = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Attachment")).ToBytes();
        var error = Assert.Throws<NotSupportedException>(() => PdfDocument.MergeToBytes(primary, plain));
        Assert.Contains("untagged", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(PdfMergeStructureMode.KeepPrimary, PdfMergeCollisionMode.RenameIncoming)]
    [InlineData(PdfMergeStructureMode.Drop, PdfMergeCollisionMode.RenameIncoming)]
    [InlineData(PdfMergeStructureMode.Combine, PdfMergeCollisionMode.KeepFirst)]
    public void MergeRejectsPoliciesThatCanDiscardTaggedFormWidgets(PdfMergeStructureMode mode, PdfMergeCollisionMode collision) {
        byte[] first = TaggedForm("first");
        byte[] second = TaggedForm("second");
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { Forms = mode, FormFieldCollisions = collision } };
        Assert.Throws<NotSupportedException>(() => PdfDocument.MergeToBytes(options, first, second));
    }

    [Fact]
    public void CombineRenamesTaggedFieldsAndKeepsEachWidgetStructureLink() {
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy {
            Forms = PdfMergeStructureMode.Combine, FormFieldCollisions = PdfMergeCollisionMode.RenameIncoming
        } };
        byte[] output = PdfDocument.MergeToBytes(options, TaggedForm("first"), TaggedForm("second"));
        var read = PdfReadDocument.Open(output);
        Assert.Equal(new[] { "Name", "Name_source2" }, read.FormFields.Select(field => field.Name).OrderBy(name => name));
        var tags = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(output).TaggedContent);
        Assert.Equal(2, tags.StructureElements.Count(element => element.StructureType == "Form"));
        var catalog = (PdfDictionary)read.Objects[read.Security.RootObjectNumber!.Value].Value;
        var root = (PdfDictionary)PdfObjectLookup.Resolve(read.Objects, catalog.Items["StructTreeRoot"])!;
        var tree = (PdfDictionary)PdfObjectLookup.Resolve(read.Objects, root.Items["ParentTree"])!;
        var nums = (PdfArray)tree.Items["Nums"];
        var keys = Enumerable.Range(0, nums.Items.Count / 2).Select(index => (int)((PdfNumber)nums.Items[index * 2]).Value).ToArray();
        foreach (var page in read.Pages) {
            var annotations = (PdfArray)PdfObjectLookup.Resolve(read.Objects, page.PageDictionary.Items["Annots"])!;
            var widget = (PdfDictionary)PdfObjectLookup.Resolve(read.Objects, Assert.Single(annotations.Items))!;
            Assert.Contains((int)widget.Get<PdfNumber>("StructParent")!.Value, keys);
        }
    }

    private static byte[] TaggedForm(string value) => PdfDocument.Create().TaggedPdfCatalogMarkers()
        .TextField("Name", value: value, style: new PdfFormFieldStyle { AlternateName = "Customer name" }).ToBytes();

    [Fact]
    public void TaggedMergeKeepsPrimaryAttachmentsWithoutRewritingTheirStructure() {
        byte[] primary = PdfDocument.Create().TaggedPdfCatalogMarkers()
            .AttachFile("notes.txt", new byte[] { 1, 2, 3 }, "text/plain")
            .H1("Primary report").ToBytes();
        byte[] output = PdfDocument.MergeToBytes(primary, Report("Incoming", "en-US"));
        var document = PdfDocument.Load(output);
        Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(document.Attachments.Extract()).Bytes);
        AssertParentTreeMatchesPages(PdfReadDocument.Open(output), 2);
    }

    [Theory]
    [InlineData("metadata", PdfMergeStructureMode.Drop, false)]
    [InlineData("metadata", PdfMergeStructureMode.Combine, false)]
    [InlineData("outlines", PdfMergeStructureMode.Drop, false)]
    [InlineData("outlines", PdfMergeStructureMode.Combine, false)]
    [InlineData("attachments", PdfMergeStructureMode.Combine, false)]
    [InlineData("attachments", PdfMergeStructureMode.Drop, true)]
    [InlineData("attachments", PdfMergeStructureMode.KeepPrimary, true)]
    public void TaggedMergeExplainsUnsupportedCatalogPoliciesBeforeAssembly(string structure, PdfMergeStructureMode mode, bool incomingAttachment) {
        byte[] first = Report("Primary", "en-US");
        byte[] second = incomingAttachment
            ? PdfDocument.Create().TaggedPdfCatalogMarkers().AttachFile("notes.txt", new byte[] { 1, 2, 3 }).H1("Incoming report").ToBytes()
            : Report("Incoming", "en-US");
        byte[] firstSnapshot = (byte[])first.Clone(), secondSnapshot = (byte[])second.Clone();
        var policy = new PdfMergePolicy();
        switch (structure) {
            case "metadata": policy.Metadata = mode; break;
            case "outlines": policy.Outlines = mode; break;
            case "attachments": policy.Attachments = mode; break;
        }
        var error = Assert.Throws<NotSupportedException>(() => PdfDocument.MergeToBytes(new PdfMergeOptions { Policy = policy }, first, second));
        Assert.Contains("Tagged PDF merge", error.Message, StringComparison.Ordinal);
        Assert.Contains(structure, error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(firstSnapshot, first);
        Assert.Equal(secondSnapshot, second);
    }

    [Fact]
    public void InterleavingRequiresAnExplicitTaggedReadingOrderPolicy() {
        byte[] source = Report("Service", "en-US");
        var error = Assert.Throws<NotSupportedException>(() => PdfPageInterleaver.Interleave(source, source));
        Assert.Contains("reading-order", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InterleavingStopsBeforeReadingSourcesAfterCancellation(bool cancelBeforeEnumeration) {
        using var stop = new CancellationTokenSource();
        int visited = 0;
        if (cancelBeforeEnumeration) stop.Cancel();
        IEnumerable<PdfInterleaveSource> Sources() {
            visited++;
            stop.Cancel();
            yield return new PdfInterleaveSource(Report("Service", "en-US"));
            throw new InvalidOperationException("Cancelled enumeration must not continue.");
        }
        Assert.Throws<OperationCanceledException>(() => PdfPageInterleaver.Interleave(Sources(),
            new PdfInterleaveOptions { CancellationToken = stop.Token }));
        Assert.Equal(cancelBeforeEnumeration ? 0 : 1, visited);
    }
}
