using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
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
                Assert.True(Enum.IsDefined(typeof(OfficeConversionLossKind), WordImageExportDiagnosticClassifier.Classify(code))));
        }

        [Theory]
        [InlineData(WordImageExportDiagnosticCodes.LimitedSmartArt, OfficeConversionLossKind.Approximation)]
        [InlineData(WordImageExportDiagnosticCodes.UnsupportedShape, OfficeConversionLossKind.Omission)]
        [InlineData(WordImageExportDiagnosticCodes.UnsupportedHeaderFooterElement, OfficeConversionLossKind.Omission)]
        [InlineData(WordImageExportDiagnosticCodes.UnsupportedHeaderElement, OfficeConversionLossKind.Omission)]
        public void WordImageExportDiagnosticClassifier_SeparatesApproximationsFromOmissions(
            string code,
            OfficeConversionLossKind expected) {
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
                WordImageTextWrapping.InLineWithText,
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
                    diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void WordDocument_StrictOmissionPolicyHandlesUnsupportedHeaderDuringMeasurement() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.HeaderDefaultOrCreate._header!.Append(
                new BookmarkStart {
                    Name = "UnsupportedHeaderMarker",
                    Id = "1"
                });
            document.AddParagraph("Body");

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                document.ExportImage(
                    OfficeImageExportFormat.Svg,
                    new WordImageExportOptions {
                        Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                    }));

            Assert.Contains(
                exception.Diagnostics,
                diagnostic =>
                    diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedHeaderElement &&
                    diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void WordDocument_StrictOmissionPolicyAllowsLimitedSmartArtFallback() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            WordSmartArt smartArt = document.AddParagraph().AddSmartArt(WordSmartArtType.BasicProcess);
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
                    diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        }
    }
}
