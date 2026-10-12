using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfPageAttributeBatchTests {
    [Theory]
    [InlineData(0, -40, 260, 420, 920)]
    [InlineData(90, -30, 270, 410, 950)]
    [InlineData(180, -20, 280, 440, 940)]
    [InlineData(270, -10, 250, 430, 930)]
    public void DisplayMarginsResolveInheritedRotationOriginAndUserUnit(int rotation,
        double left, double bottom, double right, double top) {
        byte[] source = GeometrySource(rotation);
        PdfDocument document = PdfDocument.Load(source);
        PdfDocument output = document.Pages.ApplyAttributes(PdfPageAttributeEdit.CropMargins(20, 40, 60, 80));
        PdfPageBox crop = output.Inspect().Pages[0].CropBox!;
        Assert.Equal(left, crop.Left, 8);
        Assert.Equal(bottom, crop.Bottom, 8);
        Assert.Equal(right, crop.Right, 8);
        Assert.Equal(top, crop.Top, 8);
        Assert.Equal(rotation, output.Inspect().Pages[0].RotationDegrees);
        Assert.True(output.ToBytes().AsSpan(0, source.Length).SequenceEqual(source));
        Assert.Equal(document.Inspect().Security.RevisionCount + 1, output.Inspect().Security.RevisionCount);
        // A crop changes visibility while retaining source content and page unit scale.
        Assert.Contains("Page geometry", PdfTextExtractor.ExtractAllText(output.ToBytes()));
        Assert.Equal(2D, output.Inspect().Pages[0].UserUnit);
        PdfDocument reset = output.Pages.ApplyAttributes(PdfPageAttributeEdit.ResetCrop());
        Assert.Equal(-100D, reset.Inspect().Pages[0].CropBox!.Left);
        Assert.Equal(200D, reset.Inspect().Pages[0].CropBox!.Bottom);
        Assert.Equal(500D, reset.Inspect().Pages[0].CropBox!.Right);
        Assert.Equal(1000D, reset.Inspect().Pages[0].CropBox!.Top);
    }

    [Fact]
    public void OrderedBatchUsesCurrentGeometryAndWritesOneRevision() {
        byte[] source = GeometrySource(90);
        PdfDocument document = PdfDocument.Load(source);
        PdfOperationResult<PdfDocument> result = document.Pages.ApplyAttributesResult(new[] {
            PdfPageAttributeEdit.RotateBy(90),
            PdfPageAttributeEdit.SetRotation(270),
            PdfPageAttributeEdit.RotateBy(90),
            PdfPageAttributeEdit.SetBox(PdfPageBoundaryBox.CropBox, 0, 300, 400, 900),
            PdfPageAttributeEdit.CropMargins(20, 40, 60, 80)
        });
        PdfDocument output = result.RequireValue();
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan!.ExecutionMode);
        Assert.Equal(0, output.Inspect().Pages[0].RotationDegrees);
        PdfPageBox crop = output.Inspect().Pages[0].CropBox!;
        Assert.Equal(10D, crop.Left); Assert.Equal(320D, crop.Bottom);
        Assert.Equal(370D, crop.Right); Assert.Equal(860D, crop.Top);
        Assert.Equal(document.Inspect().Security.RevisionCount + 1, output.Inspect().Security.RevisionCount);
        Assert.Equal(PdfMutationOperation.ModifyPageAttributes, output.Pipeline.Steps.Last().MutationOperation);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output.ToBytes());
        Assert.Equal(0, independent.GetPage(1).Rotation.Value);
    }

    [Fact]
    public void RelativeSelectorsReportTheRouteAndBatchFailureLeavesSourceUnchanged() {
        byte[] source = GeometrySource(90);
        PdfDocument document = PdfDocument.Load(source);
        foreach (PdfOperationResult<PdfDocument> result in new[] {
            document.Pages.RotateByResult(90, "last"),
            document.Pages.RotateByResult(90, PdfPageSelection.From(1)),
            document.Pages.RotateByResult(90, PdfPageSelector.Parse("odd"))
        }) {
            Assert.Equal(180, result.RequireValue().Inspect().Pages[0].RotationDegrees);
            Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan!.ExecutionMode);
        }
        Assert.Equal(180, document.Pages.RotateBy(90, "all").Inspect().Pages[0].RotationDegrees);
        Assert.Equal(180, document.Pages.RotateBy(90, PdfPageSelection.From(1)).Inspect().Pages[0].RotationDegrees);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Pages.ApplyAttributes(
            PdfPageAttributeEdit.RotateBy(90), PdfPageAttributeEdit.CropMargins(10000, 0, 0, 0)));
        Assert.Equal(source, document.ToBytes());
        Assert.Null(document.Pages.RotateByResult(90, "nonsense").Value);
        Assert.ThrowsAny<OperationCanceledException>(() => document.Pages.ApplyAttributesResult(
            new[] { PdfPageAttributeEdit.RotateBy(90) }, cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void CancellationAfterWritingPreventsOutputAdoptionAndSuccess() {
        byte[] source = GeometrySource(90);
        PdfDocument document = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        int stepsBefore = document.Pipeline.Steps.Count;
        Assert.ThrowsAny<OperationCanceledException>(() => document.TryMutationOperation(
            "Edit page attributes", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageAttributes,
            () => document.ApplyMutation(input => {
                byte[] output = PdfPageEditor.EditPageAttributes(input, new[] { PdfPageAttributeEdit.RotateBy(90) }, document.ReadOptions);
                cancellation.Cancel();
                return output;
            }, cancellationToken: cancellation.Token), cancellationToken: cancellation.Token));
        Assert.Equal(source, document.ToBytes());
        Assert.Equal(stepsBefore, document.Pipeline.Steps.Count);
    }

    private static byte[] GeometrySource(int rotation) {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Page geometry")).ToBytes();
        PdfReadDocument document = PdfReadDocument.Open(source);
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        int pageId = document.Pages[0].ObjectNumber;
        var page = (PdfDictionary)objects[pageId].Value;
        page.Items["MediaBox"] = Box(-100, 200, 500, 1000);
        page.Items["CropBox"] = Box(-50, 240, 450, 960);
        page.Items["UserUnit"] = new PdfNumber(2);
        page.Items.Remove("Rotate");
        int parentId = ((PdfReference)page.Items["Parent"]).ObjectNumber;
        ((PdfDictionary)objects[parentId].Value).Items["Rotate"] = new PdfNumber(rotation);
        return PdfIncrementalObjectWriter.Append(source, objects, document.Security, trailer,
            new[] { pageId, parentId }, format: PdfIncrementalXrefFormat.XrefStream);
    }

    private static PdfArray Box(params double[] values) {
        var result = new PdfArray();
        foreach (double value in values) result.Items.Add(new PdfNumber(value));
        return result;
    }
}
