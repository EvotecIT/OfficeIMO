using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class PdfRedactionWorkflowTests {
    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public async Task TextOnlyApplyAndVerifyReportTheSameRemovalCount(PdfRedactionTextSelection selection) {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf"), output = scope.PathFor("redacted.pdf");
        var background = OfficeIMO.Drawing.OfficeShape.Rectangle(300D, 60D);
        background.FillColor = OfficeIMO.Drawing.OfficeColor.Blue;
        PdfDocument.Create(pdf => pdf.Page(page => page.Canvas(canvas => {
            canvas.Shape(background, 60D, 50D);
            canvas.Text("Alpha secret Omega", 72D, 65D, 260D, 25D, fontSize: 20D);
        }))).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("secret");
        recipe.Rules[0].TextSelection = selection;
        recipe.Rules[0].ContentScope = PdfRedactionContentScope.TextOnly;
        PdfRedactionPlan nativePlan = PdfDocument.Load(input).Redactions.Search(new PdfRedactionSearchOptions {
            TextSelection = selection, ContentScope = PdfRedactionContentScope.TextOnly
        }.AddLiteral("secret"));
        Assert.Contains(nativePlan.Matches, match => match.Kind == PdfRedactionMatchKind.VectorPath);
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe
        });
        Assert.True(planned.Succeeded, planned.Summary);
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256, RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { Assert.Single(planned.Candidates).Id }
        };
        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify, InputPath = input, OutputPath = output,
            Recipe = recipe, Decisions = decisions
        });
        Assert.True(applied.Succeeded, applied.Summary);
        PdfRedactionWorkflowResult verified = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.VerifyExistingOutput, InputPath = input, OutputPath = output,
            Recipe = recipe, Decisions = decisions, ExpectedOutputSha256 = applied.Evidence!.OutputSha256
        });
        Assert.True(verified.Succeeded, verified.Summary);
        Assert.Equal(1, applied.Evidence!.VerifiedAbsentCount);
        Assert.Equal(applied.Evidence.VerifiedAbsentCount, verified.Evidence!.VerifiedAbsentCount);
        Assert.Equal(applied.Evidence.ResidualCount, verified.Evidence.ResidualCount);
        Assert.Equal(applied.Evidence.InconclusiveCount, verified.Evidence.InconclusiveCount);
    }

    [Fact]
    public async Task PreciseRecipeRoundTripPreservesNeighboursAndBindsDecisionsToSelectionPolicy() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        string output = scope.PathFor("redacted.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Alpha secret Omega")).Save(input);
        PdfRedactionRecipe recipe = JsonSerializer.Deserialize("""
            { "rules": [{ "name": "sensitive-word", "kind": "Literal", "value": "secret",
                          "textSelection": "MatchedGlyphs", "contentScope": "TextOnly" }] }
            """, PdfRedactionWorkflowJsonContext.Default.PdfRedactionRecipe)!;
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe
        });
        Assert.True(planned.Succeeded, planned.Summary);
        PdfRedactionWorkflowCandidate candidate = Assert.Single(planned.Candidates);
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256, RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { candidate.Id }
        };
        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify, InputPath = input, OutputPath = output,
            Recipe = recipe, Decisions = decisions
        });
        Assert.True(applied.Succeeded, applied.Summary);
        Assert.True(applied.Evidence?.Verified);
        string text = PdfDocument.Load(output).Reader.Text();
        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);

        recipe.Rules[0].TextSelection = PdfRedactionTextSelection.LogicalBlocks;
        PdfRedactionWorkflowResult stale = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify, InputPath = input,
            OutputPath = scope.PathFor("stale.pdf"), Recipe = recipe, Decisions = decisions
        });
        Assert.False(stale.Succeeded);
        Assert.Contains(stale.Diagnostics, diagnostic => diagnostic.Message.Contains("different recipe revision", StringComparison.Ordinal));
        Assert.False(File.Exists(scope.PathFor("stale.pdf")));
    }

    [Fact]
    public async Task PreciseWorkflowBlocksUnsafeNativeSearchBeforePublishingCandidates() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        string evidence = scope.PathFor("plan.json");
        PdfDocument.Create().Paragraph(paragraph => paragraph.CharacterSpacing(-3).Text("AB")).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("A");
        recipe.Rules[0].TextSelection = PdfRedactionTextSelection.MatchedGlyphs;
        PdfRedactionWorkflowResult result = await new OfficeWorkflowRunner().RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe, EvidencePath = evidence
        });
        Assert.False(result.Succeeded);
        Assert.Empty(result.Candidates);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("RedactionSearchUnselectedTextIntersection", StringComparison.Ordinal));
        Assert.False(File.Exists(evidence));
        Assert.Contains("AB", PdfDocument.Load(input).Reader.Text(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PdfRedactionDetectionMode.NativeAndOcr, PdfRedactionRuleKind.Literal)]
    [InlineData(PdfRedactionDetectionMode.NativeOnly, PdfRedactionRuleKind.LogicalKind)]
    public async Task PreciseRecipeRejectsUnsupportedDetectionOrRuleCombinations(PdfRedactionDetectionMode detection, PdfRedactionRuleKind kind) {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("secret")).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("secret");
        recipe.DetectionMode = detection;
        recipe.Rules[0].Kind = kind;
        recipe.Rules[0].TextSelection = PdfRedactionTextSelection.MatchedGlyphs;
        PdfRedactionWorkflowResult result = await new OfficeWorkflowRunner().RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe
        });
        Assert.False(result.Succeeded);
        Assert.Empty(result.Candidates);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RedactionWorkflowFailed" &&
            diagnostic.Details?["exceptionType"] == nameof(ArgumentException));
    }
}
