using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class PdfRedactionWorkflowTests {
    [Fact]
    public async Task DirectoryBatchRejectsManifestInsidePhysicalInputRoot() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        Directory.CreateDirectory(inputRoot);
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(Path.Combine(inputRoot, "one.pdf"));

        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            EvidenceRoot = scope.PathFor("evidence"),
            ManifestPath = Path.Combine(inputRoot, "batch.json"),
            Recipe = CreateRecipe("secret")
        }));
    }

    [Fact]
    public async Task ExplicitBatchInputRejectsPhysicalSymlinkEscapeWhenSupported() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        string outsideRoot = scope.PathFor("outside");
        Directory.CreateDirectory(inputRoot);
        Directory.CreateDirectory(outsideRoot);
        string outsidePdf = Path.Combine(outsideRoot, "outside.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(outsidePdf);
        string link = Path.Combine(inputRoot, "linked.pdf");
        try {
            File.CreateSymbolicLink(link, outsidePdf);
        } catch (Exception exception) when (exception is UnauthorizedAccessException or PlatformNotSupportedException or IOException) {
            return;
        }

        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            InputPaths = { "linked.pdf" },
            EvidenceRoot = scope.PathFor("evidence"),
            ManifestPath = scope.PathFor("batch.json"),
            Recipe = CreateRecipe("secret")
        }));
    }

    [Fact]
    public async Task RecursiveBatchDiscoveryDoesNotFollowDirectorySymlinksWhenSupported() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        string outsideRoot = scope.PathFor("outside");
        Directory.CreateDirectory(inputRoot);
        Directory.CreateDirectory(outsideRoot);
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(Path.Combine(outsideRoot, "outside.pdf"));
        try {
            Directory.CreateSymbolicLink(Path.Combine(inputRoot, "linked"), outsideRoot);
        } catch (Exception exception) when (exception is UnauthorizedAccessException or PlatformNotSupportedException or IOException) {
            return;
        }

        PdfRedactionBatchResult result = await new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            EvidenceRoot = scope.PathFor("evidence"),
            ManifestPath = scope.PathFor("batch.json"),
            Recipe = CreateRecipe("secret")
        });

        Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
        Assert.Empty(result.Items);
    }

    [Fact]
    public async Task DirectoryBatchStopsAtConfiguredItemLimitBeforePlanning() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        Directory.CreateDirectory(inputRoot);
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(Path.Combine(inputRoot, "one.pdf"));
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(Path.Combine(inputRoot, "two.pdf"));

        await Assert.ThrowsAsync<InvalidOperationException>(() => new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            EvidenceRoot = scope.PathFor("evidence"),
            ManifestPath = scope.PathFor("batch.json"),
            Recipe = CreateRecipe("secret"),
            Limits = new PdfRedactionWorkflowLimits { MaximumBatchItems = 1 }
        }));
    }

    [Fact]
    public async Task ContinuePerItemBatchCannotReplaceReviewedDecisionWithManifest() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        string outputRoot = scope.PathFor("output");
        string evidenceRoot = scope.PathFor("evidence");
        string decisionsRoot = scope.PathFor("decisions");
        Directory.CreateDirectory(inputRoot);
        Directory.CreateDirectory(decisionsRoot);
        string input = Path.Combine(inputRoot, "one.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("secret");
        PdfRedactionWorkflowResult planned = await new OfficeWorkflowRunner().RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe
        });
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { Assert.Single(planned.Candidates).Id }
        };
        string decisionPath = Path.Combine(decisionsRoot, "one.decisions.json");
        string originalDecision = JsonSerializer.Serialize(decisions, PdfRedactionWorkflowJsonContext.Default.PdfRedactionDecisionManifest);
        await File.WriteAllTextAsync(decisionPath, originalDecision);

        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputRoot = inputRoot,
            OutputRoot = outputRoot,
            EvidenceRoot = evidenceRoot,
            DecisionsRoot = decisionsRoot,
            ManifestPath = decisionPath,
            PublicationPolicy = PdfRedactionBatchPublicationPolicy.ContinuePerItem,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            Recipe = recipe
        }));

        Assert.Equal(originalDecision, await File.ReadAllTextAsync(decisionPath));
        Assert.False(File.Exists(Path.Combine(outputRoot, "one.redacted.pdf")));
        Assert.False(File.Exists(Path.Combine(evidenceRoot, "one.redaction.json")));
    }
}
