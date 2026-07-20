using System.IO;
using System.Reflection;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordImageExportTests {
        [Fact]
        public void WordImageExportDiagnosticClassifier_CoversEveryPublishedCode() {
            string[] codes = typeof(WordImageExportDiagnosticCodes)
                .GetFields(BindingFlags.Public | BindingFlags.Static)
                .Where(field => field.IsLiteral && field.FieldType == typeof(string))
                .Select(field => Assert.IsType<string>(field.GetRawConstantValue()))
                .ToArray();

            Assert.NotEmpty(codes);
            Assert.All(codes, code =>
                Assert.True(Enum.IsDefined(typeof(OfficeImageExportLossKind), WordImageExportDiagnosticClassifier.Classify(code))));
        }

        [Theory]
        [InlineData(WordImageExportDiagnosticCodes.LimitedSmartArt, OfficeImageExportLossKind.Approximation)]
        [InlineData(WordImageExportDiagnosticCodes.UnsupportedShape, OfficeImageExportLossKind.Omission)]
        [InlineData(WordImageExportDiagnosticCodes.UnsupportedHeaderElement, OfficeImageExportLossKind.Omission)]
        public void WordImageExportDiagnosticClassifier_SeparatesApproximationsFromOmissions(
            string code,
            OfficeImageExportLossKind expected) {
            Assert.Equal(expected, WordImageExportDiagnosticClassifier.Classify(code));
        }

        [Fact]
        public void WordImageExportDiagnosticClassifier_RejectsUnknownCode() {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                WordImageExportDiagnosticClassifier.Classify("limited-word-never-published"));
        }

        [Fact]
        public void WordDocument_StrictOmissionPolicyRejectsSkippedVisualContent() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourcePng = CreateSolidPng(420, 420, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage image = document.AddParagraph().InsertImage(
                imageStream,
                "strict-rotated-inline.png",
                420,
                420,
                WrapTextImage.InLineWithText,
                "Strict rotated inline marker");
            image.Rotation = 45;

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                document.ExportImage(
                    OfficeImageExportFormat.Svg,
                    new WordImageExportOptions {
                        Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                    }));

            Assert.Contains(
                exception.Diagnostics,
                diagnostic =>
                    diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedImage &&
                    diagnostic.LossKind == OfficeImageExportLossKind.Omission);
        }

        [Fact]
        public void WordDocument_StrictOmissionPolicyAllowsLimitedSmartArtFallback() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            WordSmartArt smartArt = document.AddParagraph().AddSmartArt(SmartArtType.BasicProcess);
            while (smartArt.NodeCount < 3) {
                smartArt.AddNode("Node " + smartArt.NodeCount);
            }
            smartArt.ReplaceTexts("Plan", "Build", "Ship");

            OfficeImageExportResult result = document.ExportImage(
                OfficeImageExportFormat.Svg,
                new WordImageExportOptions {
                    Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                });

            Assert.Contains(
                result.Diagnostics,
                diagnostic =>
                    diagnostic.Code == WordImageExportDiagnosticCodes.LimitedSmartArt &&
                    diagnostic.LossKind == OfficeImageExportLossKind.Approximation);
        }
    }
}
