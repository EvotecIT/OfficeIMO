using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageAttributeIncrementalTests {
    [Theory]
    [InlineData("OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf")]
    [InlineData("OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-pdfa1b-text.pdf")]
    public void RotateResult_AppendsOnlySelectedPageAndReopensIndependentProducer(string fixture) {
        byte[] source = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), fixture));
        PdfDocument original = PdfDocument.Load(source);
        PdfDocumentInfo before = original.Inspect();
        int pageObjectNumber = PdfReadDocument.Open(source).Pages[0].ObjectNumber;
        Assert.Equal(!before.HasTaggedContent, original.PlanMutation(PdfMutationOperation.ModifyPageTree).CanExecute);

        PdfOperationResult<PdfDocument> result = original.Pages.RotateResult(90, "1");
        PdfDocument rotated = result.RequireValue();
        byte[] output = rotated.ToBytes();

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan!.ExecutionMode);
        Assert.Contains(result.MutationPlan.Diagnostics, text => text.Contains("because"));
        Assert.True(result.MutationPlan.AppendOnlyReport.CanAppendPageAttributes);
        AssertRevision(source, output, pageObjectNumber);
        PdfDocumentInfo after = PdfDocument.Load(output).Inspect();
        Assert.Equal(90, after.Pages[0].RotationDegrees);
        Assert.Equal(before.HasTaggedContent, after.HasTaggedContent);
        Assert.Equal(before.PageCount, after.PageCount);
        Assert.Equal(pageObjectNumber, PdfReadDocument.Open(output).Pages[0].ObjectNumber);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Equal(before.PageCount, independent.NumberOfPages);
        Assert.Equal(90, independent.GetPage(1).Rotation.Value);
        PdfOperationResult<PdfDocument> selectorResult = original.Pages.RotateResult(90, PdfPageSelector.Parse("1"));
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, selectorResult.MutationPlan!.ExecutionMode);
        Assert.Equal(90, selectorResult.RequireValue().Inspect().Pages[0].RotationDegrees);
        Assert.Equal(PdfMutationOperation.RotatePages, rotated.Pipeline.Steps.Last().MutationOperation);

        PdfDocument insertedPage = PdfDocument.Create().Paragraph(p => p.Text("Inserted"));
        Assert.Equal(before.PageCount + 1, original.Pages.Insert(1, insertedPage).Inspect().PageCount);
        if (before.HasTaggedContent)
            Assert.Throws<PdfMutationBlockedException>(() => original.Pages.CropAndTranslate(0, 0, 100, 100, 1));
    }

    [Theory]
    [InlineData(PdfPageBoundaryBox.MediaBox)]
    [InlineData(PdfPageBoundaryBox.CropBox)]
    [InlineData(PdfPageBoundaryBox.TrimBox)]
    [InlineData(PdfPageBoundaryBox.BleedBox)]
    [InlineData(PdfPageBoundaryBox.ArtBox)]
    public void SetPageBoxResult_AppendsBoundaryWithoutChangingContent(PdfPageBoundaryBox box) {
        byte[] source = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf"));
        PdfDocument original = PdfDocument.Load(source);
        int pageObjectNumber = PdfReadDocument.Open(source).Pages[0].ObjectNumber;
        PdfOperationResult<PdfDocument> result = original.Pages.SetPageBoxResult(box, 10, 20, 200, 300, PdfPageSelection.From(1));
        byte[] output = result.RequireValue().ToBytes();

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan!.ExecutionMode);
        Assert.Equal(PdfMutationOperation.ModifyPageAttributes, result.RequireValue().Pipeline.Steps.Last().MutationOperation);
        Assert.Equal("SetPageBox", result.RequireValue().Pipeline.Steps.Last().Operation);
        AssertRevision(source, output, pageObjectNumber);
        var (objects, _) = PdfSyntax.ParseObjects(output);
        var page = Assert.IsType<PdfDictionary>(objects[pageObjectNumber].Value);
        var bounds = Assert.IsType<PdfArray>(page.Items[box.ToString()]);
        Assert.Equal(new double[] { 10, 20, 200, 300 }, bounds.Items.Cast<PdfNumber>().Select(number => number.Value));
    }

    [Fact]
    public void InheritedRotation_OverridesOnlyRequestedPageAndKeepsAbsoluteAngleContract() {
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(c => c.Paragraph(p => p.Text("First"))));
            compose.Page(page => page.Content(c => c.Paragraph(p => p.Text("Second"))));
        }).ToBytes();
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        PdfReadDocument read = PdfReadDocument.Open(source);
        var first = Assert.IsType<PdfDictionary>(objects[read.Pages[0].ObjectNumber].Value);
        var parentReference = Assert.IsType<PdfReference>(first.Items["Parent"]);
        var parent = Assert.IsType<PdfDictionary>(objects[parentReference.ObjectNumber].Value);
        parent.Items["Rotate"] = new PdfNumber(270);
        parent.Items["CropBox"] = Box(10, 20, 400, 500);
        source = PdfIncrementalObjectWriter.Append(source, objects, PdfSyntax.ReadDocumentSecurityInfo(source), trailer,
            new[] { parentReference.ObjectNumber }, format: PdfIncrementalXrefFormat.XrefStream);
        Assert.All(PdfInspector.Inspect(source).Pages, page => Assert.Equal(270, page.RotationDegrees));

        byte[] rotated = PdfDocument.Load(source).Pages.Rotate(-270, 1).ToBytes();
        AssertRevision(source, rotated, read.Pages[0].ObjectNumber);
        PdfDocumentInfo after = PdfInspector.Inspect(rotated);
        Assert.Equal(90, after.Pages[0].RotationDegrees);
        Assert.Equal(270, after.Pages[1].RotationDegrees);
        byte[] turned = PdfDocument.Load(rotated).Pages.RotateBy(90, 1, 2).ToBytes();
        AssertRevision(rotated, turned, read.Pages[0].ObjectNumber, read.Pages[1].ObjectNumber);
        Assert.Equal(new[] { 180, 0 }, PdfInspector.Inspect(turned).Pages.Select(page => page.RotationDegrees));
        byte[] cropped = PdfDocument.Load(rotated).Pages.SetCropBox(30, 40, 300, 400, 1).ToBytes();
        var (afterObjects, _) = PdfSyntax.ParseObjects(cropped);
        Assert.False(Assert.IsType<PdfDictionary>(afterObjects[read.Pages[1].ObjectNumber].Value).Items.ContainsKey("CropBox"));
        Assert.Equal(10, Assert.IsType<PdfNumber>(Assert.IsType<PdfArray>(
            Assert.IsType<PdfDictionary>(afterObjects[parentReference.ObjectNumber].Value).Items["CropBox"]).Items[0]).Value);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Certification_BlocksRotationAndBoxesAtEveryPermissionLevel(int permissionLevel) {
        byte[] source = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel);
        PdfDocument document = PdfDocument.Load(source);
        PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.ModifyPageAttributes);
        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.False(plan.AppendOnlyReport.CanAppendPageAttributes);
        Assert.Contains("AppendOnly.DocMDP", plan.BlockerCodes);
        Assert.Throws<PdfMutationBlockedException>(() => document.Pages.Rotate(90, 1));
        Assert.Throws<PdfMutationBlockedException>(() => document.Pages.SetCropBox(10, 10, 100, 100, 1));
    }

    [Fact]
    public void PageAttributes_DoNotPersistUnrelatedParentRepairs() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Repaired parent")).ToBytes();
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        PdfReadDocument read = PdfReadDocument.Open(source);
        int pageNumber = read.Pages[0].ObjectNumber;
        int wrongParent = read.Security.RootObjectNumber!.Value;
        var page = Assert.IsType<PdfDictionary>(objects[pageNumber].Value);
        page.Items["Parent"] = new PdfReference(wrongParent, 0);
        source = PdfIncrementalObjectWriter.Append(source, objects, read.Security, trailer, new[] { pageNumber }, format: PdfIncrementalXrefFormat.XrefStream);
        Assert.NotEqual(wrongParent, Assert.IsType<PdfReference>(
            Assert.IsType<PdfDictionary>(PdfReadDocument.Open(source).Objects[pageNumber].Value).Items["Parent"]).ObjectNumber);

        foreach (PdfDocument output in new[] {
            PdfDocument.Load(source).Pages.Rotate(90, 1),
            PdfDocument.Load(source).Pages.SetCropBox(10, 20, 100, 150, 1)
        }) {
            var (afterObjects, _) = PdfSyntax.ParseObjects(output.ToBytes());
            var updated = Assert.IsType<PdfDictionary>(afterObjects[pageNumber].Value);
            Assert.Equal(wrongParent, Assert.IsType<PdfReference>(updated.Items["Parent"]).ObjectNumber);
            Assert.Equal(page.Items.Count + 1, updated.Items.Count);
            AssertRevision(source, output.ToBytes(), pageNumber);
        }
    }

    [Fact]
    public void OrdinarySource_KeepsFullRewriteAndReportsIt() {
        PdfDocument document = PdfDocument.Create().Paragraph(p => p.Text("Ordinary PDF"));
        PdfOperationResult<PdfDocument> result = document.Pages.RotateResult(180, "1");
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan!.ExecutionMode);
        Assert.Equal(180, result.RequireValue().Inspect().Pages[0].RotationDegrees);
        Assert.Equal(270, document.Pages.RotateBy(-90, 1).Inspect().Pages[0].RotationDegrees);
        PdfDocument boxed = document.Pages.SetCropBox(10, 20, 200, 300, 1);
        Assert.Equal(PdfMutationOperation.ModifyPageAttributes, boxed.Pipeline.Steps.Last().MutationOperation);
    }

    private static void AssertRevision(byte[] source, byte[] output, params int[] pageObjectNumbers) {
        Assert.True(output.AsSpan(0, source.Length).SequenceEqual(source));
        PdfDocumentSecurityInfo before = PdfSyntax.ReadDocumentSecurityInfo(source);
        PdfDocumentSecurityInfo after = PdfSyntax.ReadDocumentSecurityInfo(output);
        Assert.Equal(before.RevisionCount + 1, after.RevisionCount);
        Assert.Equal(before.LastStartXrefOffset, after.Revisions.Last().PreviousXrefOffset);
        string appended = Encoding.ASCII.GetString(output, source.Length, output.Length - source.Length);
        int[] appendedObjects = Regex.Matches(appended, @"(?m)^(\d+) \d+ obj\s*$")
            .Cast<Match>().Select(match => int.Parse(match.Groups[1].Value)).ToArray();
        Assert.All(pageObjectNumbers, number => Assert.Contains(number, appendedObjects));
        bool stream = output[after.LastStartXrefOffset!.Value] != (byte)'x';
        Assert.Equal(pageObjectNumbers.Length + (stream ? 1 : 0), appendedObjects.Length);
        Assert.Equal(source[before.LastStartXrefOffset!.Value] != (byte)'x', stream);
    }

    private static PdfArray Box(params double[] coordinates) {
        var box = new PdfArray();
        foreach (double coordinate in coordinates) box.Items.Add(new PdfNumber(coordinate));
        return box;
    }
}
