using System.Text.Json;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfRedactionWorkflowTests {
    [Fact]
    public async Task PlanApplyAndVerifyPublishesPrivacySafeEvidence() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        string planPath = scope.PathFor("plan.json");
        string output = scope.PathFor("redacted.pdf");
        string evidencePath = scope.PathFor("evidence.json");
        const string sensitive = "WorkflowSecret-4917";
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive)).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        var runner = new OfficeWorkflowRunner();

        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Id = "private-request-correlation",
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            EvidencePath = planPath,
            Recipe = recipe
        });

        Assert.True(planned.Succeeded, planned.Summary);
        PdfRedactionWorkflowCandidate candidate = Assert.Single(planned.Candidates);
        string persistedPlan = await File.ReadAllTextAsync(planPath);
        Assert.DoesNotContain(sensitive, persistedPlan, StringComparison.Ordinal);
        Assert.DoesNotContain("private-request-correlation", persistedPlan, StringComparison.Ordinal);
        Assert.DoesNotContain(scope.DirectoryPath, persistedPlan, StringComparison.OrdinalIgnoreCase);
        using (JsonDocument planJson = JsonDocument.Parse(persistedPlan)) {
            Assert.Equal("officeimo.pdf.redaction.plan.v1", planJson.RootElement.GetProperty("schema").GetString());
        }
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { candidate.Id }
        };

        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            EvidencePath = evidencePath,
            Recipe = recipe,
            Decisions = decisions
        });

        Assert.True(applied.Succeeded, string.Join(Environment.NewLine, applied.Diagnostics.Select(static diagnostic => diagnostic.Code + ": " + diagnostic.Message + " [" + string.Join(",", diagnostic.Details.Select(static pair => pair.Key + "=" + pair.Value)) + "]")));
        Assert.True(applied.Evidence?.Verified);
        Assert.DoesNotContain(sensitive, PdfDocument.Load(output).Reader.Text(), StringComparison.Ordinal);
        string evidence = await File.ReadAllTextAsync(evidencePath);
        Assert.DoesNotContain(sensitive, evidence, StringComparison.Ordinal);
        Assert.DoesNotContain("extractedText", evidence, StringComparison.OrdinalIgnoreCase);
        using JsonDocument json = JsonDocument.Parse(evidence);
        Assert.Equal("officeimo.pdf.redaction.result.v1", json.RootElement.GetProperty("schema").GetString());
    }

    [Fact]
    public async Task ApplyRejectsIncompleteOrStaleReviewWithoutPublishing() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        string output = scope.PathFor("redacted.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("ReviewSecret-220")).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("ReviewSecret-220");
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest { Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe });

        var decisions = new PdfRedactionDecisionManifest { SourceSha256 = planned.SourceSha256, RecipeSha256 = planned.RecipeSha256 };
        PdfRedactionWorkflowResult result = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            Decisions = decisions
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(File.Exists(output));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("explicitly approve or reject", StringComparison.Ordinal));
    }

    [Fact]
    public async Task SignedPolicyIsExplicitAndNeverPublishes() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("signed.pdf");
        string output = scope.PathFor("redacted.pdf");
        byte[] unsigned = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("signed content")).ToBytes();
        byte[] signedMarker = unsigned.Concat(System.Text.Encoding.ASCII.GetBytes("\n% /Type /Sig\n")).ToArray();
        await File.WriteAllBytesAsync(input, signedMarker);

        PdfRedactionWorkflowResult result = await new OfficeWorkflowRunner().RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = CreateRecipe("signed")
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("rejects signed sources", StringComparison.Ordinal));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task SignedSourceCanBeRedactedIntoAResignedDerivativeWithIndependentEvidence() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("signed-source.pdf");
        string output = scope.PathFor("signed-redacted.pdf");
        const string sensitive = "SignedWorkflowSecret-884";
        byte[] unsigned = PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive)).ToBytes();
        PdfExternalSignaturePreparation sourcePreparation = PdfIncrementalUpdater.PrepareExternalSignature(unsigned, new PdfExternalSignatureOptions {
            FieldName = "SourceSignature",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(sourcePreparation, Enumerable.Repeat((byte)0x33, 128).ToArray());
        await File.WriteAllBytesAsync(input, signed);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        recipe.SignaturePolicy = PdfRedactionSignaturePolicy.CreateAndSignDerivative;
        var runner = new OfficeWorkflowRunner();

        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe
        });
        PdfRedactionWorkflowCandidate candidate = Assert.Single(planned.Candidates);
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { candidate.Id }
        };

        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            Decisions = decisions,
            OutputSigner = new FixedSignatureSigner(),
            OutputSignatureOptions = new PdfExternalSignatureOptions { FieldName = "DerivativeSignature", ReservedSignatureContentsBytes = 512 },
            ExternalValidators = { new HeaderExternalValidator() }
        });

        Assert.True(applied.Succeeded, string.Join(Environment.NewLine, applied.Diagnostics.Select(static diagnostic => diagnostic.Message)));
        Assert.Equal(1, applied.Evidence?.SourceSignatureCount);
        Assert.Equal(1, applied.Evidence?.OutputSignatureCount);
        Assert.Equal(nameof(PdfRedactionSignaturePolicy.CreateAndSignDerivative), applied.Evidence?.SignaturePolicy);
        Assert.Equal("fixed-test-signer", applied.Evidence?.OutputSigner);
        Assert.Contains("header-validator", applied.Evidence!.ExternalValidators);
        PdfDocument derivative = PdfDocument.Load(output);
        Assert.True(derivative.Security.ValidateSignatures().IsStructurallyValid);
        Assert.DoesNotContain(sensitive, derivative.Reader.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task IndependentValidatorRejectionPreventsPublication() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("validator-source.pdf");
        string output = scope.PathFor("validator-redacted.pdf");
        const string sensitive = "ValidatorSecret-219";
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive)).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe
        });
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { Assert.Single(planned.Candidates).Id }
        };

        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            Decisions = decisions,
            ExternalValidators = { new RejectingExternalValidator() }
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, applied.Status);
        Assert.False(File.Exists(output));
        Assert.Contains(applied.Diagnostics, static diagnostic => diagnostic.Message.Contains("independent validators rejected", StringComparison.Ordinal));
    }

    [Fact]
    public async Task CancellationStopsRunningIndependentValidatorWithoutPublication() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("validator-cancellation-source.pdf");
        string output = scope.PathFor("validator-cancellation-redacted.pdf");
        const string sensitive = "ValidatorCancellationSecret-220";
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive)).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe
        });
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { Assert.Single(planned.Candidates).Id }
        };
        using var validator = new BlockingExternalValidator();
        using var cancellation = new CancellationTokenSource();

        Task<PdfRedactionWorkflowResult> running = runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            Decisions = decisions,
            ExternalValidators = { validator }
        }, cancellationToken: cancellation.Token);

        Assert.True(validator.WaitUntilStarted(TimeSpan.FromSeconds(10)), "The independent validator did not start.");
        cancellation.Cancel();
        PdfRedactionWorkflowResult result = await running;

        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task EncryptedWorkflowRequiresExplicitPolicyAndCanReencryptVerifiedOutput() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("protected.pdf");
        string output = scope.PathFor("protected-redacted.pdf");
        const string ownerPassword = "owner-redaction-1";
        const string outputPassword = "output-redaction-2";
        const string sensitive = "ProtectedSecret-712";
        byte[] encrypted = PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive))
            .Security.Encrypt(new PdfStandardEncryptionOptions("reader-redaction-1") { OwnerPassword = ownerPassword }).Pdf;
        await File.WriteAllBytesAsync(input, encrypted);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        var runner = new OfficeWorkflowRunner();

        PdfRedactionWorkflowResult rejected = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe,
            OwnerPassword = ownerPassword
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, rejected.Status);

        recipe.EncryptedDocumentPolicy = PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt;
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe,
            OwnerPassword = ownerPassword,
            OutputEncryption = new PdfStandardEncryptionOptions(outputPassword)
        });
        PdfRedactionWorkflowCandidate candidate = Assert.Single(planned.Candidates);
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { candidate.Id }
        };

        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            Decisions = decisions,
            OwnerPassword = ownerPassword,
            OutputEncryption = new PdfStandardEncryptionOptions(outputPassword)
        });

        Assert.True(applied.Succeeded, string.Join(Environment.NewLine, applied.Diagnostics.Select(static diagnostic => diagnostic.Message)));
        PdfDocument outputDocument = PdfDocument.Load(output, new PdfLoadOptions { Password = outputPassword });
        Assert.True(outputDocument.Inspect().Security.HasEncryption);
        Assert.DoesNotContain(sensitive, outputDocument.Reader.Text(), StringComparison.Ordinal);
        Assert.Equal(PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt.ToString(), applied.Evidence?.EncryptionPolicy);
    }

    [Fact]
    public async Task EncryptedWorkflowRequiresOwnerAuthorization() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("protected.pdf");
        const string userPassword = "reader-preserve-1";
        const string ownerPassword = "owner-preserve-2";
        const string sensitive = "PreservedSecret-810";
        byte[] encrypted = PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive))
            .Security.Encrypt(new PdfStandardEncryptionOptions(userPassword) { OwnerPassword = ownerPassword }).Pdf;
        await File.WriteAllBytesAsync(input, encrypted);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        recipe.EncryptedDocumentPolicy = PdfRedactionEncryptedDocumentPolicy.Decrypt;
        var runner = new OfficeWorkflowRunner();

        PdfRedactionWorkflowResult userRejected = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe,
            OwnerPassword = userPassword
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, userRejected.Status);
        Assert.Contains(userRejected.Diagnostics, diagnostic => diagnostic.Message.Contains("owner password", StringComparison.OrdinalIgnoreCase));

        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe,
            OwnerPassword = ownerPassword
        });
        Assert.True(planned.Succeeded, string.Join(Environment.NewLine, planned.Diagnostics.Select(static diagnostic => diagnostic.Message)));
        Assert.Single(planned.Candidates);
    }

    [Fact]
    public async Task BatchConflictPublishesNothingAndPreservesExistingDestination() {
        using var scope = new RedactionTestDirectory();
        var runner = new OfficeWorkflowRunner();
        var requests = new List<PdfRedactionWorkflowRequest>();
        string firstOutput = scope.PathFor("first-redacted.pdf");
        string secondOutput = scope.PathFor("second-redacted.pdf");
        byte[] sentinel = { 1, 2, 3, 4 };
        await File.WriteAllBytesAsync(secondOutput, sentinel);
        for (int index = 0; index < 2; index++) {
            string secret = "BatchSecret-" + index;
            string input = scope.PathFor("source-" + index + ".pdf");
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text(secret)).Save(input);
            PdfRedactionRecipe recipe = CreateRecipe(secret);
            PdfRedactionWorkflowResult plan = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest { Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe });
            requests.Add(new PdfRedactionWorkflowRequest {
                Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
                InputPath = input,
                OutputPath = index == 0 ? firstOutput : secondOutput,
                Recipe = recipe,
                Decisions = new PdfRedactionDecisionManifest {
                    SourceSha256 = plan.SourceSha256,
                    RecipeSha256 = plan.RecipeSha256,
                    ApprovedCandidateIds = { Assert.Single(plan.Candidates).Id }
                }
            });
        }

        PdfRedactionBatchResult result = await runner.RunRedactionBatchAsync(requests);

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(result.PublishedAtomically);
        Assert.False(File.Exists(firstOutput));
        Assert.Equal(sentinel, await File.ReadAllBytesAsync(secondOutput));
        Assert.All(result.Items, item => {
            Assert.Equal(OfficeWorkflowStatus.Failed, item.Status);
            Assert.Null(item.OutputPath);
            Assert.Null(item.EvidencePath);
            Assert.Null(item.Evidence);
        });
    }

    [Fact]
    public async Task RejectedCandidatesStillHonorDecryptAndReencryptPolicy() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("protected.pdf");
        string output = scope.PathFor("protected-copy.pdf");
        const string ownerPassword = "zero-owner-password";
        const string outputPassword = "zero-output-password";
        byte[] encrypted = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("RejectedProtectedCandidate"))
            .Security.Encrypt(new PdfStandardEncryptionOptions("zero-reader-password") { OwnerPassword = ownerPassword }).Pdf;
        await File.WriteAllBytesAsync(input, encrypted);
        PdfRedactionRecipe recipe = CreateRecipe("RejectedProtectedCandidate");
        recipe.EncryptedDocumentPolicy = PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt;
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult plan = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe, OwnerPassword = ownerPassword
        });
        PdfRedactionWorkflowCandidate candidate = Assert.Single(plan.Candidates);

        PdfRedactionWorkflowResult result = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            OwnerPassword = ownerPassword,
            OutputEncryption = new PdfStandardEncryptionOptions(outputPassword),
            Decisions = new PdfRedactionDecisionManifest {
                SourceSha256 = plan.SourceSha256,
                RecipeSha256 = plan.RecipeSha256,
                RejectedCandidateIds = { candidate.Id }
            }
        });

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Diagnostics.Select(static diagnostic => diagnostic.Message)));
        Assert.True(PdfDocument.Load(output, new PdfLoadOptions { Password = outputPassword }).Inspect().Security.HasEncryption);
        Assert.Equal(PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt.ToString(), result.Evidence?.EncryptionPolicy);
        Assert.NotEqual(encrypted, await File.ReadAllBytesAsync(output));

        PdfRedactionWorkflowResult verified = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.VerifyExistingOutput,
            InputPath = input,
            OutputPath = output,
            Recipe = recipe,
            OwnerPassword = ownerPassword,
            OutputEncryption = new PdfStandardEncryptionOptions(outputPassword),
            ExpectedOutputSha256 = result.Evidence!.OutputSha256,
            Decisions = new PdfRedactionDecisionManifest {
                SourceSha256 = plan.SourceSha256,
                RecipeSha256 = plan.RecipeSha256,
                RejectedCandidateIds = { candidate.Id }
            }
        });
        Assert.True(verified.Succeeded, string.Join(Environment.NewLine, verified.Diagnostics.Select(static diagnostic => diagnostic.Message)));

        string wrongPolicyOutput = scope.PathFor("protected-copy-aes128.pdf");
        byte[] decrypted = PdfDocument.Load(output, new PdfLoadOptions { Password = outputPassword })
            .Security.Decrypt(outputPassword).Pdf;
        byte[] wrongPolicyBytes = PdfDocument.Load(decrypted).Security.Encrypt(new PdfStandardEncryptionOptions(outputPassword) {
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        }).Pdf;
        await File.WriteAllBytesAsync(wrongPolicyOutput, wrongPolicyBytes);

        PdfRedactionWorkflowResult wrongPolicy = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.VerifyExistingOutput,
            InputPath = input,
            OutputPath = wrongPolicyOutput,
            Recipe = recipe,
            OwnerPassword = ownerPassword,
            OutputEncryption = new PdfStandardEncryptionOptions(outputPassword),
            ExpectedOutputSha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(wrongPolicyBytes)).ToLowerInvariant(),
            Decisions = new PdfRedactionDecisionManifest {
                SourceSha256 = plan.SourceSha256,
                RecipeSha256 = plan.RecipeSha256,
                RejectedCandidateIds = { candidate.Id }
            }
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, wrongPolicy.Status);
        Assert.Contains(wrongPolicy.Diagnostics, static diagnostic =>
            diagnostic.Message.Contains("encryption algorithm", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task ProgressCallbacksCannotRetargetValidatedPublication() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        string output = scope.PathFor("redacted.pdf");
        const string sensitive = "SnapshotSecret-491";
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text(sensitive)).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult plan = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest { Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe });
        var request = new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            Recipe = recipe,
            Decisions = new PdfRedactionDecisionManifest {
                SourceSha256 = plan.SourceSha256,
                RecipeSha256 = plan.RecipeSha256,
                ApprovedCandidateIds = { Assert.Single(plan.Candidates).Id }
            }
        };
        var progress = new MutatingProgress<OfficeWorkflowProgress>(_ => {
            request.OutputPath = input;
            request.Recipe.CleanupScope = PdfRedactionCleanupScope.None;
        });

        PdfRedactionWorkflowResult result = await runner.RunRedactionAsync(request, progress);

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Diagnostics.Select(static diagnostic => diagnostic.Message)));
        Assert.Contains(sensitive, PdfDocument.Load(input).Reader.Text(), StringComparison.Ordinal);
        Assert.DoesNotContain(sensitive, PdfDocument.Load(output).Reader.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task RichRegionProducesOneAtomicCandidateForAllNormalizedAreas() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("group region")) .Save(input);
        var recipe = new PdfRedactionRecipe();
        recipe.Regions.Add(new PdfRedactionRecipeRegion {
            Kind = PdfRedactionRegionKind.Group,
            PageNumber = 1,
            Areas = {
                new PdfRedactionRecipeRegion { Kind = PdfRedactionRegionKind.Rectangle, PageNumber = 1, X = 20, Y = 20, Width = 30, Height = 10 },
                new PdfRedactionRecipeRegion { Kind = PdfRedactionRegionKind.Rectangle, PageNumber = 1, X = 100, Y = 100, Width = 40, Height = 20 }
            }
        });

        PdfRedactionWorkflowResult plan = await new OfficeWorkflowRunner().RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe
        });

        PdfRedactionWorkflowCandidate candidate = Assert.Single(plan.Candidates);
        Assert.Equal(2, candidate.Areas.Count);
    }

    [Fact]
    public async Task AtomicBatchPreparedByteBudgetFailsBeforePublishing() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        string evidence = scope.PathFor("plan.json");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("budget target")).Save(input);
        var request = new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            EvidencePath = evidence,
            Recipe = CreateRecipe("budget target"),
            Limits = new PdfRedactionWorkflowLimits { MaximumBatchPreparedBytes = 1024 }
        };

        PdfRedactionBatchResult result = await new OfficeWorkflowRunner().RunRedactionBatchAsync(new[] { request });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(result.PublishedAtomically);
        Assert.False(File.Exists(evidence));
    }

    [Fact]
    public async Task OcrAssistedApplyRunsPostRewriteOcrAndPersistsOnlySafeMetadata() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("ocr-source.pdf");
        string output = scope.PathFor("ocr-redacted.pdf");
        string evidencePath = scope.PathFor("ocr-evidence.json");
        const string sensitive = "Ocr Secret";
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("unrelated native content")).Save(input);
        int invocation = 0;
        var engine = new DelegateOcrEngine(
            "workflow-fixture",
            (_, _) => {
                invocation++;
                return Task.FromResult(invocation < 3
                    ? new OcrResult {
                        Provider = "fixture-provider",
                        Model = "fixture-model",
                        Language = "en",
                        Spans = new[] {
                            OcrWord("Ocr", 200, 300, 30, 14, 0.98),
                            OcrWord("Secret", 235, 300, 50, 14, 0.94)
                        }
                    }
                    : new OcrResult { Provider = "fixture-provider", Model = "fixture-model", Language = "en" });
            },
            new OcrEngineCapabilities { SupportsWordSpans = true, SupportsConfidence = true });
        PdfRedactionRecipe recipe = CreateRecipe(sensitive);
        recipe.DetectionMode = PdfRedactionDetectionMode.OcrOnly;
        var runner = new OfficeWorkflowRunner();

        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe,
            OcrEngine = engine
        });
        PdfRedactionWorkflowCandidate candidate = Assert.Single(planned.Candidates);
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { candidate.Id }
        };

        PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = output,
            EvidencePath = evidencePath,
            Recipe = recipe,
            Decisions = decisions,
            OcrEngine = engine
        });

        Assert.True(applied.Succeeded, string.Join(Environment.NewLine, applied.Diagnostics.Select(static diagnostic => diagnostic.Message)));
        Assert.Equal(3, invocation);
        Assert.True(applied.Evidence?.OcrUsed);
        Assert.True(applied.Evidence?.OcrPostVerificationPerformed);
        Assert.Equal(0, applied.Evidence?.OcrResidualCandidateCount);
        Assert.Contains("fixture-provider", applied.Evidence!.OcrProviders);
        string evidence = await File.ReadAllTextAsync(evidencePath);
        Assert.DoesNotContain(sensitive, evidence, StringComparison.Ordinal);
        Assert.DoesNotContain("recognized", evidence, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task OcrProviderExceptionTextIsNotExposedByPrivacySafeResult() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("ocr-source.pdf");
        const string leaked = "ProviderReturnedSensitiveText-991";
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("page")).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("secret");
        recipe.DetectionMode = PdfRedactionDetectionMode.OcrOnly;
        var engine = new DelegateOcrEngine("failing-provider", (_, _) => throw new InvalidOperationException(leaked));

        PdfRedactionWorkflowResult result = await new OfficeWorkflowRunner().RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.PlanOnly,
            InputPath = input,
            Recipe = recipe,
            OcrEngine = engine
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.DoesNotContain(leaked, result.Summary, StringComparison.Ordinal);
        Assert.All(result.Diagnostics, diagnostic => Assert.DoesNotContain(leaked, diagnostic.Message, StringComparison.Ordinal));
    }

    [Fact]
    public async Task ApplyRejectsInPlaceOutputAndEvidenceCollisions() {
        using var scope = new RedactionTestDirectory();
        string input = scope.PathFor("source.pdf");
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text("CollisionSecret")).Save(input);
        PdfRedactionRecipe recipe = CreateRecipe("CollisionSecret");
        var runner = new OfficeWorkflowRunner();
        PdfRedactionWorkflowResult planned = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest { Mode = PdfRedactionWorkflowMode.PlanOnly, InputPath = input, Recipe = recipe });
        var decisions = new PdfRedactionDecisionManifest {
            SourceSha256 = planned.SourceSha256,
            RecipeSha256 = planned.RecipeSha256,
            ApprovedCandidateIds = { Assert.Single(planned.Candidates).Id }
        };

        PdfRedactionWorkflowResult result = await runner.RunRedactionAsync(new PdfRedactionWorkflowRequest {
            Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
            InputPath = input,
            OutputPath = input,
            EvidencePath = input,
            Recipe = recipe,
            Decisions = decisions,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("CollisionSecret", PdfDocument.Load(input).Reader.Text(), StringComparison.Ordinal);
    }

    private static PdfRedactionRecipe CreateRecipe(string value) {
        var recipe = new PdfRedactionRecipe();
        recipe.Rules.Add(new PdfRedactionRule { Kind = PdfRedactionRuleKind.Literal, Value = value });
        return recipe;
    }

    private static OcrTextSpan OcrWord(string text, double x, double y, double width, double height, double confidence) => new() {
        Level = OcrTextSpanLevel.Word,
        Text = text,
        Confidence = confidence,
        CoordinateUnit = OcrCoordinateUnit.Points,
        Region = new OcrRegion { X = x, Y = y, Width = width, Height = height }
    };

    private sealed class RedactionTestDirectory : IDisposable {
        private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OfficeIMO.RedactionWorkflow.Tests", Guid.NewGuid().ToString("N"));
        internal RedactionTestDirectory() => Directory.CreateDirectory(_path);
        internal string DirectoryPath => _path;
        internal string PathFor(string name) => System.IO.Path.Combine(_path, name);
        public void Dispose() { if (Directory.Exists(_path)) Directory.Delete(_path, recursive: true); }
    }

    private sealed class MutatingProgress<T>(Action<T> callback) : IProgress<T> {
        public void Report(T value) => callback(value);
    }

    private sealed class FixedSignatureSigner : IPdfExternalSigner {
        public string Name => "fixed-test-signer";
        public byte[] Sign(PdfExternalSignatureRequest request) => Enumerable.Repeat((byte)0x44, 128).ToArray();
    }

    private sealed class HeaderExternalValidator : IPdfRedactionCancellationAwareExternalValidator {
        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) =>
            new("header-validator", redactedPdf.AsSpan().StartsWith("%PDF-"u8));

        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Validate(redactedPdf);
        }
    }

    private sealed class RejectingExternalValidator : IPdfRedactionCancellationAwareExternalValidator {
        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) =>
            new("rejecting-validator", false, "fixture rejection");

        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Validate(redactedPdf);
        }
    }

    private sealed class BlockingExternalValidator : IPdfRedactionCancellationAwareExternalValidator, IDisposable {
        private readonly ManualResetEventSlim _started = new(false);

        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) =>
            throw new InvalidOperationException("The workflow must use the cancellation-aware validator contract.");

        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf, CancellationToken cancellationToken) {
            _started.Set();
            cancellationToken.WaitHandle.WaitOne();
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfRedactionExternalValidationResult("blocking-validator", true);
        }

        internal bool WaitUntilStarted(TimeSpan timeout) => _started.Wait(timeout);

        public void Dispose() => _started.Dispose();
    }
}
