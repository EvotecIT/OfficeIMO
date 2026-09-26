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
    public async Task BatchRejectsInputReplacedWithOutsideSymlinkAfterDiscovery() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        Directory.CreateDirectory(inputRoot);
        string input = Path.Combine(inputRoot, "one.pdf");
        string outside = scope.PathFor("outside.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(input);
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("outside secret")).Save(outside);
        string probe = Path.Combine(inputRoot, "probe.pdf");
        try {
            File.CreateSymbolicLink(probe, outside);
            File.Delete(probe);
        } catch (Exception exception) when (exception is UnauthorizedAccessException or PlatformNotSupportedException or IOException) {
            return;
        }
        bool replaced = false;
        var progress = new BatchProgress(update => {
            if (replaced || update.Stage != "validate") return;
            File.Delete(input);
            File.CreateSymbolicLink(input, outside);
            replaced = true;
        });

        PdfRedactionBatchResult result = await new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            EvidenceRoot = scope.PathFor("evidence"),
            ManifestPath = scope.PathFor("batch.json"),
            Recipe = CreateRecipe("secret")
        }, progress);

        Assert.True(replaced);
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(File.Exists(Path.Combine(scope.PathFor("evidence"), "one.redaction.json")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BatchPublicationCannotFollowDirectoryReplacedAfterDiscovery(bool replaceManifestDirectory) {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        string evidenceRoot = scope.PathFor("evidence");
        string manifestRoot = scope.PathFor("manifest");
        string outsideRoot = scope.PathFor("outside");
        string evidenceChild = Path.Combine(evidenceRoot, "nested");
        Directory.CreateDirectory(Path.Combine(inputRoot, "nested"));
        Directory.CreateDirectory(evidenceChild);
        Directory.CreateDirectory(manifestRoot);
        Directory.CreateDirectory(outsideRoot);
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret"))
            .Save(Path.Combine(inputRoot, "nested", "one.pdf"));

        string target = replaceManifestDirectory ? manifestRoot : evidenceChild;
        string probe = Path.Combine(scope.PathFor("probe"), "link");
        Directory.CreateDirectory(Path.GetDirectoryName(probe)!);
        try {
            Directory.CreateSymbolicLink(probe, outsideRoot);
            Directory.Delete(probe);
        } catch (Exception exception) when (exception is UnauthorizedAccessException or PlatformNotSupportedException or IOException) {
            return;
        }
        bool replaced = false;
        var progress = new BatchProgress(update => {
            if (replaced || update.Stage != "validate") return;
            Directory.Delete(target);
            Directory.CreateSymbolicLink(target, outsideRoot);
            replaced = true;
        });

        PdfRedactionBatchResult result = await new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            EvidenceRoot = evidenceRoot,
            ManifestPath = Path.Combine(manifestRoot, "batch.json"),
            RecurseSubdirectories = true,
            Recipe = CreateRecipe("secret")
        }, progress);

        Assert.True(replaced);
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(File.Exists(Path.Combine(outsideRoot, "one.redaction.json")));
        Assert.False(File.Exists(Path.Combine(outsideRoot, "batch.json")));
        Assert.False(File.Exists(Path.Combine(evidenceChild, "one.redaction.json")));
    }

    [Fact]
    public async Task BatchReplacePublishesEvidenceAndManifestWithoutLeavingRollbackFiles() {
        using var scope = new RedactionTestDirectory();
        string inputRoot = scope.PathFor("input");
        string evidenceRoot = scope.PathFor("evidence");
        Directory.CreateDirectory(inputRoot);
        Directory.CreateDirectory(evidenceRoot);
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret"))
            .Save(Path.Combine(inputRoot, "one.pdf"));
        string evidence = Path.Combine(evidenceRoot, "one.redaction.json");
        string manifest = scope.PathFor("batch.json");
        await File.WriteAllTextAsync(evidence, "old evidence");
        await File.WriteAllTextAsync(manifest, "old manifest");

        PdfRedactionBatchResult result = await new OfficeWorkflowRunner().RunRedactionBatchAsync(new PdfRedactionBatchRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputRoot = inputRoot,
            EvidenceRoot = evidenceRoot,
            ManifestPath = manifest,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            Recipe = CreateRecipe("secret")
        });

        Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
        Assert.NotEqual("old evidence", await File.ReadAllTextAsync(evidence));
        Assert.NotEqual("old manifest", await File.ReadAllTextAsync(manifest));
        Assert.Empty(Directory.EnumerateFiles(scope.DirectoryPath, "*.rollback", SearchOption.AllDirectories));
    }

    private sealed class BatchProgress(Action<OfficeWorkflowProgress> report) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => report(value);
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
