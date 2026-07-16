using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void FeatureReportPreflight_AllowsCleanDocumentWorkflows() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Clean document");

            WordFeatureReport report = document.InspectFeatures();

            foreach (WordPreflightCapability capability in Enum.GetValues(typeof(WordPreflightCapability))) {
                Assert.True(report.Can(capability));
                Assert.Same(report, report.EnsureCan(capability));
                Assert.Empty(report.GetCapabilityDiagnostics(capability));
                Assert.Empty(report.GetRepairHints(capability));
            }
            Assert.Contains("## Capability Preflight", report.ToMarkdown(), StringComparison.Ordinal);
        }

        [Fact]
        public void FeatureReportPreflight_DistinguishesSignedReadRenderAndMutationWorkflows() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Signed metadata");
            document.ApplicationProperties.DigitalSignature = new DigitalSignature();

            WordFeatureReport report = document.InspectFeatures();

            Assert.True(report.CanReadDocumentContent);
            Assert.True(report.CanRenderFixedLayout);
            Assert.False(report.CanEditDocumentContent);
            Assert.False(report.CanEditDocumentStructure);
            Assert.False(report.CanBindTemplate);
            Assert.False(report.CanSavePackageRoundTrip);
            Assert.Contains(report.GetCapabilityDiagnostics(WordPreflightCapability.SavePackageRoundTrip),
                diagnostic => diagnostic.Contains("Digital signatures", StringComparison.Ordinal));
            WordPreflightRepairHint hint = Assert.Single(
                report.GetRepairHints(WordPreflightCapability.SavePackageRoundTrip));
            Assert.Equal("Digital signatures", hint.FeatureName);
            Assert.Contains("unsigned copy", hint.Action, StringComparison.OrdinalIgnoreCase);
            Assert.Throws<InvalidOperationException>(() =>
                report.EnsureCan(WordPreflightCapability.EditDocumentContent));
        }

        [Fact]
        public void FeatureReportPreflight_BlocksFixedLayoutForUnmaterializedAlternativeContent() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Native content");
            document.AddEmbeddedFragment("<html><body><p>Imported</p></body></html>",
                WordAlternativeFormatImportPartType.Html);

            WordFeatureReport report = document.InspectFeatures();

            Assert.True(report.CanReadDocumentContent);
            Assert.False(report.CanRenderFixedLayout);
            Assert.Contains(report.GetCapabilityDiagnostics(WordPreflightCapability.RenderFixedLayout),
                diagnostic => diagnostic.Contains("Alternative format imports", StringComparison.Ordinal));
            WordPreflightRepairHint hint = Assert.Single(
                report.GetRepairHints(WordPreflightCapability.RenderFixedLayout));
            Assert.Contains("Materialize", hint.Action, StringComparison.OrdinalIgnoreCase);
            string markdown = report.ToMarkdown();
            Assert.Contains("| RenderFixedLayout | No |", markdown, StringComparison.Ordinal);
            Assert.Contains("## Repair And Routing Hints", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void FeatureReportPreflight_AllowsContentEditsButBlocksStructureForPreserveOnlyParts() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Preserve-only metadata");
            CustomXmlPart part = document._wordprocessingDocument.MainDocumentPart!
                .AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<root><value>1</value></root>"))) {
                part.FeedData(stream);
            }

            WordFeatureReport report = document.InspectFeatures();

            Assert.True(report.CanReadDocumentContent);
            Assert.True(report.CanEditDocumentContent);
            Assert.True(report.CanRenderFixedLayout);
            Assert.True(report.CanSavePackageRoundTrip);
            Assert.False(report.CanEditDocumentStructure);
            Assert.False(report.CanBindTemplate);
            Assert.Contains(report.GetCapabilityDiagnostics(WordPreflightCapability.EditDocumentStructure),
                diagnostic => diagnostic.Contains("Custom XML parts", StringComparison.Ordinal));
            Assert.Contains(report.GetRepairHints(WordPreflightCapability.EditDocumentStructure),
                hint => hint.FeatureName == "Custom XML parts");
        }

        [Fact]
        public void FeatureReportPreflight_RejectsUnknownCapabilityValues() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Capability validation");
            WordFeatureReport report = document.InspectFeatures();
            WordPreflightCapability unknown = (WordPreflightCapability)int.MaxValue;

            Assert.Throws<ArgumentOutOfRangeException>(() => report.Can(unknown));
            Assert.Throws<ArgumentOutOfRangeException>(() => report.GetCapabilityDiagnostics(unknown));
            Assert.Throws<ArgumentOutOfRangeException>(() => report.GetRepairHints(unknown));
        }
    }
}
