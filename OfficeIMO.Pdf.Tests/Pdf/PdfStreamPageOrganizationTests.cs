using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStreamPageOrganizationTests {
    [Fact]
    public void CompressedSourceSupportsPageOrganizationThroughAnExplicitNormalizingRewrite() {
        byte[] source = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf"));
        PdfDocument document = PdfDocument.Load(source);
        PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.ModifyPageTree);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, plan.ExecutionMode);
        Assert.Contains("Input.CrossReferenceAndObjectStreamsWillBeNormalized", plan.Warnings);
        PdfOperationResult<PdfDocument> duplicated = document.Pages.DuplicateResult("1");
        PdfDocument copies = duplicated.RequireValue();
        Assert.Equal(document.Inspect().PageCount + 1, copies.Inspect().PageCount);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, duplicated.MutationPlan!.ExecutionMode);
        PdfDocument reversed = copies.Pages.Reorder(Enumerable.Range(1, copies.Inspect().PageCount).Reverse().ToArray());
        PdfDocument deleted = reversed.Pages.Delete(1);
        Assert.Equal(document.Inspect().PageCount, deleted.Inspect().PageCount);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(deleted.ToBytes());
        Assert.Equal(deleted.Inspect().PageCount, independent.NumberOfPages);
        using var original = UglyToad.PdfPig.PdfDocument.Open(source);
        using var duplicatedReader = UglyToad.PdfPig.PdfDocument.Open(copies.ToBytes());
        Assert.Equal(original.GetPage(1).Text, duplicatedReader.GetPage(1).Text);
        Assert.Equal(original.GetPage(1).Text, duplicatedReader.GetPage(2).Text);
    }
}
