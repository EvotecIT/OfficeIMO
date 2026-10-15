using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStudioConsolidationRegressionTests {
    [Fact]
    public void AssemblyPermissionAllowsRotationButNotBoundaryEditsOrMixedBatches() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Assembly permission"))
            .Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.AssembleDocument
            }).Pdf;
        var reader = PdfDocument.Load(source, new PdfLoadOptions { Password = "reader" });
        Assert.True(reader.PlanMutation(PdfMutationOperation.RotatePages).CanExecute);
        Assert.False(reader.PlanMutation(PdfMutationOperation.ModifyPageAttributes).CanExecute);
        var ownerOptions = new PdfLoadOptions { Password = "owner" };
        Assert.Equal(90, PdfReadDocument.Open(reader.Pages.RotateBy(90).ToBytes(), ownerOptions).Pages[0].GetRotationDegrees());
        Assert.Equal(90, PdfReadDocument.Open(reader.Pages.ApplyAttributes(PdfPageAttributeEdit.SetRotation(90)).ToBytes(), ownerOptions).Pages[0].GetRotationDegrees());
        Assert.Throws<PdfMutationBlockedException>(() => reader.Pages.SetCropBox(10, 10, 200, 200));
        Assert.Throws<PdfMutationBlockedException>(() => reader.Pages.ApplyAttributes(
            PdfPageAttributeEdit.RotateBy(90), PdfPageAttributeEdit.CropMargins(10, 10, 10, 10)));
        var owner = PdfDocument.Load(source, new PdfLoadOptions { Password = "owner" });
        Assert.Equal(10, owner.Pages.SetCropBox(10, 10, 200, 200).Inspect().Pages[0].CropBox!.Left);
        Assert.Equal(source, reader.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageAttributeRewritePreservesOpaqueCatalogEntriesAndMetadata(bool crop) {
        byte[] original = PdfDocument.Create().Paragraph(p => p.Text("Keep document state")).ToBytes();
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var custom = new PdfDictionary(); custom.Items["Payload"] = new PdfStringObj("preserve me");
            catalog.Items["StudioExtension"] = custom;
            int infoNumber = security.InfoObjectNumber ?? objects.Keys.Max() + 1;
            if (!objects.TryGetValue(infoNumber, out var info)) {
                info = new PdfIndirectObject(infoNumber, 0, new PdfDictionary());
                objects.Add(infoNumber, info);
            }
            ((PdfDictionary)info.Value).Items["Title"] = new PdfStringObj("Preserved title");
            return infoNumber;
        });
        var document = PdfDocument.Load(source);
        byte[] output = (crop ? document.Pages.SetCropBox(10, 10, 200, 200) : document.Pages.RotateBy(90)).ToBytes();
        var read = PdfReadDocument.Open(output);
        var catalog = (PdfDictionary)read.Objects[read.Security.RootObjectNumber!.Value].Value;
        var custom = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(read.Objects, catalog.Items["StudioExtension"]));
        Assert.Equal("preserve me", Assert.IsType<PdfStringObj>(custom.Items["Payload"]).Value);
        Assert.Equal("Preserved title", PdfDocument.Load(output).Inspect().Metadata.Title);
        Assert.Contains("Keep document state", PdfDocument.Load(output).Reader.Text());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void UntaggedPageImportUsesTheAdvertisedCatalogPreservingRoute(int position) {
        var target = PdfDocument.Create().Paragraph(p => p.Text("Existing page"))
            .JavaScript.AddOrReplace("Initialize", "this.zoom = 125;").ToDocument();
        var incoming = PdfDocument.Create().Paragraph(p => p.Text("Imported page"));
        Assert.True(target.PlanMutation(PdfMutationOperation.ImportPages).CanExecute);
        var output = position switch {
            3 => PdfDocument.Load(PdfPageImporter.AppendPageRanges(target.ToBytes(), incoming.ToBytes(), PdfPageRange.From(1, 1))),
            4 => PdfDocument.Load(PdfPageImporter.PrependPageRanges(target.ToBytes(), incoming.ToBytes(), PdfPageRange.From(1, 1))),
            5 => PdfDocument.Load(PdfPageImporter.InsertPageRanges(target.ToBytes(), incoming.ToBytes(), 2, PdfPageRange.From(1, 1))),
            _ => target.Pages.Insert(position, incoming, new PdfPageImportOptions())
        };
        Assert.Equal(2, output.Inspect().PageCount);
        Assert.Equal("this.zoom = 125;", Assert.Single(output.JavaScript.List()).Script);
        Assert.Contains("Imported page", output.Reader.Text());
        Assert.Equal(1, target.Inspect().PageCount);
    }
}
