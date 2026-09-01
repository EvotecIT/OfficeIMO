using System.Diagnostics;
using OfficeIMO.Pdf;

namespace OfficeIMO.RealWorldCorpus;

internal static class PdfCorpusProbe {
    internal static CorpusPdfEvidence Probe(byte[] snapshot) {
        var evidence = new CorpusPdfEvidence();
        var stages = new List<CorpusPdfStageResult>();
        PdfDocument? document = null;
        PdfDocumentInfo? documentInfo = null;

        Measure(stages, "inspect", () => {
            document = PdfDocument.Load(snapshot);
            documentInfo = document.Inspect();
            evidence.PageCount = documentInfo.PageCount;
        });
        if (document is null) {
            evidence.Stages = stages.AsReadOnly();
            return evidence;
        }

        Measure(stages, "logical-semantics", () => {
            PdfDocumentReadResult logical = document.Read();
            evidence.ParagraphCount = logical.Paragraphs.Count;
            evidence.TableCount = logical.Tables.Count;
            evidence.CrossPageParagraphCount = logical.GetParagraphContinuationGroups().Count(group => group.SpansPages);
            evidence.CrossPageTableCount = logical.GetTableContinuationGroups().Count(group => group.SpansPages);
        });
        Measure(stages, "fonts", () => {
            PdfFontInventory fonts = document.Resources.Fonts();
            evidence.FontCount = fonts.FontCount;
            evidence.EmbeddedFontCount = fonts.EmbeddedFontCount;
            evidence.ParsedEmbeddedOpenTypeFontCount = fonts.Fonts.Count(static font => font.EmbeddedOpenTypeInfo is not null);
            evidence.MissingToUnicodeFontCount = fonts.MissingToUnicodeFontCount;
        });
        Measure(stages, "images", () => {
            evidence.ImageCount = document.Images.Extract().Count;
            evidence.ImagePlacementCount = document.Images.Placements().Count;
        });
        Measure(stages, "managed-first-page-render", () => {
            if (documentInfo is null || evidence.PageCount == 0) {
                throw new InvalidDataException("Document inspection did not expose a renderable PDF page.");
            }
            IReadOnlyList<PdfPageRenderResult> renders = document.Render.Pages(
                "1",
                new PdfPageRenderOptions {
                    Format = PdfPageRenderFormat.Svg,
                    MaxPages = 1,
                    MaxOutputBytesPerPage = 8L * 1024L * 1024L,
                    MaxTotalOutputBytes = 8L * 1024L * 1024L,
                    ContinueOnError = true
                });
            evidence.RenderAttemptedPages = renders.Count;
            evidence.RenderSucceededPages = renders.Count(static render => render.Succeeded);
            evidence.RenderDiagnosticCodes = renders
                .SelectMany(static render => render.CapabilityDiagnostics)
                .Select(static diagnostic => diagnostic.Code)
                .Where(static code => !string.IsNullOrWhiteSpace(code))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(static code => code, StringComparer.Ordinal)
                .ToArray();
            if (renders.Count != 1 || !renders[0].Succeeded) {
                throw new InvalidDataException("The managed first-page PDF render did not produce an artifact.");
            }
        });
        Measure(stages, "mutation-portfolio", () => {
            PdfMutationPortfolioReport portfolio = document.AssessMutations();
            evidence.MutationPlanCount = portfolio.Plans.Count;
            evidence.FullRewriteMutationPlanCount = portfolio.Plans.Count(static plan => plan.ExecutionMode == PdfMutationExecutionMode.FullRewrite);
            evidence.AppendOnlyMutationPlanCount = portfolio.Plans.Count(static plan => plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly);
            evidence.BlockedMutationPlanCount = portfolio.BlockedPlans.Count;
            evidence.MutationPlanModes = portfolio.Plans.ToDictionary(
                static plan => plan.Operation.ToString(),
                static plan => plan.ExecutionMode.ToString(),
                StringComparer.Ordinal);
        });
        Measure(stages, "declared-compliance-claims", () => {
            PdfDeclaredComplianceClaimsReport claims = document.AssessDeclaredComplianceClaims();
            evidence.DeclaredComplianceClaimCount = claims.Claims.Count;
            evidence.RecognizedComplianceClaimCount = claims.RecognizedClaims.Count;
            evidence.ClaimableComplianceClaimCount = claims.Claims.Count(static claim => claim.CanClaimConformance);
            evidence.DeclaredComplianceClaimStatuses = claims.Claims
                .Select(static claim => claim.Declaration + ":" + claim.Status)
                .ToArray();
        });

        evidence.Stages = stages.AsReadOnly();
        return evidence;
    }

    private static void Measure(List<CorpusPdfStageResult> stages, string name, Action action) {
        Stopwatch stopwatch = Stopwatch.StartNew();
        try {
            action();
            stages.Add(new CorpusPdfStageResult {
                Name = name,
                Succeeded = true,
                DurationMilliseconds = stopwatch.ElapsedMilliseconds
            });
        } catch (Exception exception) {
            stages.Add(new CorpusPdfStageResult {
                Name = name,
                Succeeded = false,
                DurationMilliseconds = stopwatch.ElapsedMilliseconds,
                ExceptionType = exception.GetType().FullName ?? exception.GetType().Name
            });
        }
    }
}
