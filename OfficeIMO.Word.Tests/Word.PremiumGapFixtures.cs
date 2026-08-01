using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class WordPremiumGapFixtureTests {
        [Fact]
        public void Test_WordPremiumGapFixtureManifest_DefinesProofContracts() {
            string manifestPath = Path.Combine(AppContext.BaseDirectory, "Documents", "Word", "PremiumGaps", "premium-gap-fixtures.xml");

            Assert.True(File.Exists(manifestPath), $"Missing premium-gap fixture manifest: {manifestPath}");

            XDocument manifest = XDocument.Load(manifestPath);
            XElement root = Assert.IsType<XElement>(manifest.Root);

            Assert.Equal("PremiumGapFixtures", root.Name.LocalName);
            Assert.Equal("1", RequiredAttribute(root, "schemaVersion"));
            Assert.Contains("OfficeIMO.Word", RequiredAttribute(root, "scope"), StringComparison.Ordinal);
            Assert.Contains("excluding PDF and legacy DOC", RequiredAttribute(root, "scope"), StringComparison.Ordinal);

            List<XElement> scenarios = root.Elements("Scenario").ToList();
            Assert.NotEmpty(scenarios);

            var expectedWorkstreams = new HashSet<string>(StringComparer.Ordinal) {
                "DigitalSignatures",
                "ReviewRedline",
                "DocumentComparison",
                "FieldEvaluation",
                "TemplateMailMerge",
                "RealWorldCorpus"
            };

            HashSet<string> actualWorkstreams = scenarios
                .Select(scenario => RequiredAttribute(scenario, "workstream"))
                .ToHashSet(StringComparer.Ordinal);

            Assert.True(
                expectedWorkstreams.SetEquals(actualWorkstreams),
                "Premium-gap manifest workstreams changed. Expected: "
                    + string.Join(", ", expectedWorkstreams.OrderBy(value => value, StringComparer.Ordinal))
                    + ". Actual: "
                    + string.Join(", ", actualWorkstreams.OrderBy(value => value, StringComparer.Ordinal)));

            var ids = new HashSet<string>(StringComparer.Ordinal);
            var allowedStatuses = new HashSet<string>(StringComparer.Ordinal) { "planned", "partially-covered", "covered" };
            var allowedFixtureStatuses = new HashSet<string>(StringComparer.Ordinal) { "needed", "available", "generated" };
            var allowedExpectedBehaviors = new HashSet<string>(StringComparer.Ordinal) {
                "inspect-preserve-warn",
                "read-report",
                "compare-report",
                "redline-document",
                "evaluate-refresh-diagnose",
                "preflight-merge-diagnose"
            };

            foreach (XElement scenario in scenarios) {
                string id = RequiredAttribute(scenario, "id");
                Assert.Matches("^[a-z0-9]+(?:-[a-z0-9]+)*$", id);
                Assert.True(ids.Add(id), $"Duplicate premium-gap scenario id: {id}");

                Assert.Contains(RequiredAttribute(scenario, "status"), allowedStatuses);
                Assert.Contains(RequiredAttribute(scenario, "fixtureStatus"), allowedFixtureStatuses);
                Assert.Contains(RequiredAttribute(scenario, "expectedBehavior"), allowedExpectedBehaviors);
                Assert.StartsWith("Documents/Word/PremiumGaps/", RequiredAttribute(scenario, "sourceDocument"), StringComparison.Ordinal);
                Assert.False(string.IsNullOrWhiteSpace(RequiredAttribute(scenario, "featureFamily")));

                List<string> evidence = scenario.Elements("Evidence")
                    .Select(element => element.Value.Trim())
                    .Where(value => value.Length > 0)
                    .ToList();
                Assert.NotEmpty(evidence);

                List<string> validationCommands = scenario.Elements("Validation")
                    .Select(element => element.Value.Trim())
                    .Where(value => value.Length > 0)
                    .ToList();
                Assert.NotEmpty(validationCommands);
                Assert.All(validationCommands, command =>
                    Assert.True(
                        command.Contains("dotnet test", StringComparison.OrdinalIgnoreCase)
                        || command.Contains("dotnet run", StringComparison.OrdinalIgnoreCase),
                        "Validation command must be a concrete dotnet test or dotnet run command: " + command));
                Assert.All(validationCommands, command =>
                    Assert.DoesNotContain("OfficeIMO.Tests/OfficeIMO.Tests.csproj", command, StringComparison.OrdinalIgnoreCase));
            }

            string repositoryRoot = LocateRepositoryRootForPremiumGapTests();
            string roadmapPath = Path.Combine(repositoryRoot, "Docs", "ROADMAP.md");
            string roadmap = File.ReadAllText(roadmapPath);
            Assert.Contains("## Word", roadmap, StringComparison.Ordinal);
            Assert.Contains("Add VBA macro-project signing as a separate explicit capability", roadmap, StringComparison.Ordinal);

            string compatibilityPath = Path.Combine(repositoryRoot, "OfficeIMO.Word", "COMPATIBILITY.md");
            string compatibility = File.ReadAllText(compatibilityPath);
            Assert.Contains("Cross-platform OPC XML-signature creation", compatibility, StringComparison.Ordinal);
            Assert.Contains("relationship-transform", compatibility, StringComparison.Ordinal);
            Assert.Contains("RFC 3161", compatibility, StringComparison.Ordinal);
            Assert.Contains("VBA macro-project signing", compatibility, StringComparison.Ordinal);
            Assert.Contains("stable limitation codes", compatibility, StringComparison.Ordinal);
            Assert.Contains("WordPageNumberBasis.ExplicitBreakEstimate", compatibility, StringComparison.Ordinal);

            string fixtureManifestPath = Path.Combine(
                repositoryRoot,
                "OfficeIMO.TestAssets",
                "Documents",
                "Word",
                "PremiumGaps",
                "premium-gap-fixtures.xml");
            Assert.True(File.Exists(fixtureManifestPath), $"Missing premium-gap fixture manifest: {fixtureManifestPath}");

            string matrixPath = Path.Combine(repositoryRoot, "Docs", "officeimo.word-template-mail-merge-scenarios.md");
            Assert.True(File.Exists(matrixPath), $"Missing template/mail-merge scenario matrix: {matrixPath}");

            string matrix = File.ReadAllText(matrixPath);
            foreach (string requiredScenario in new[] {
                "Merge fields",
                "Conditional blocks",
                "Repeated table rows",
                "Grouped table rows",
                "Repeated body blocks",
                "Nested regions",
                "Section regions",
                "Headers and footers",
                "Table cells",
                "Content controls",
                "Template diagnostics",
                "Test_MailMerge_ComplexSplitRunFieldsPreserveResultFormattingWhenKeepingFields",
                "Test_MailMerge_NestedRegionsPreserveTableCellFieldFormatting",
                "Test_MailMerge_RepeatingBlockRegionsPreserveSectionBreakProperties",
                "Test_MailMerge_ConditionalBlocksCanKeepOrRemoveSectionRegions",
                "Test_MailMerge_ContentControlFormFillPreservesTextRunFormatting",
                "Invoice",
                "Grouped table report",
                "MailMergeGroupedTableWorkflow.docx",
                "Proposal",
                "Review letter",
                "Header/footer approval package",
                "MailMergeHeaderFooterWorkflow.docx",
                "Form fill"
            }) {
                Assert.Contains(requiredScenario, matrix, StringComparison.Ordinal);
            }

            Assert.Contains("--word-mail-merge-workflows", matrix, StringComparison.Ordinal);
            Assert.Contains("WordMailMerge.PreflightTemplate", matrix, StringComparison.Ordinal);
            Assert.Contains("Word-native record controls", matrix, StringComparison.Ordinal);
            Assert.Contains("`NEXT`, `NEXTIF`, `SKIPIF`, `MERGEREC`, and `MERGESEQ`", matrix, StringComparison.Ordinal);
            Assert.Contains("`UnsupportedMailMergeControlField`", matrix, StringComparison.Ordinal);
            Assert.Contains("`CanBindTemplate` to `false`", matrix, StringComparison.Ordinal);
            Assert.Contains("Test_MailMerge_PreflightTemplateReportsUnsupportedWordNativeRecordControlFields", matrix, StringComparison.Ordinal);

            Assert.Contains("unknown-document-feature-preflight", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("feature-report.md", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("--word-review-reports", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("--word-comparison-reports", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("--word-signature-preflight", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("WordSignatureValidationReport", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("ValidateSignatures", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("imported-related-part-list-of-figures.docx", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("Test_TableOfContent_RefreshListOfFiguresSupportsImportedRelatedPartCaptionFixture", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("imported-note-part-list-of-figures.docx", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("Test_TableOfContent_RefreshListOfFiguresSupportsImportedNotePartCaptionFixture", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("imported-related-part-index.docx", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
            Assert.Contains("Test_TableOfContent_RefreshIndexSupportsImportedRelatedPartIndexFixture", manifest.ToString(SaveOptions.DisableFormatting), StringComparison.Ordinal);
        }

        private static string RequiredAttribute(XElement element, string name) {
            string? value = (string?)element.Attribute(name);
            Assert.False(string.IsNullOrWhiteSpace(value), $"Missing {name} on {element.Name.LocalName}.");
            return value!;
        }

        private static string LocateRepositoryRootForPremiumGapTests() {
            DirectoryInfo? directory = new(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            throw new InvalidOperationException("Unable to locate the OfficeIMO repository root.");
        }
    }
}
