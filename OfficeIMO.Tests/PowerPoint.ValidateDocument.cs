using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointValidateDocument {
        [Fact]
        public void Test_ValidateDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (var presentation = PowerPointPresentation.Create(filePath)) {
                var errors = presentation.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                Assert.True(presentation.DocumentIsValid);
                presentation.Save();
            }

            using (var presentation = PowerPointPresentation.Open(filePath)) {
                var errors = presentation.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                Assert.True(presentation.DocumentIsValid);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void DocumentValidationErrors_IsCachedUntilInvalidated() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            int factoryCalls = 0;
            Func<FileFormatVersions, OpenXmlValidator> originalFactory = PowerPointPresentation.ValidatorFactory;

            try {
                PowerPointPresentation.ValidatorFactory = version => {
                    factoryCalls++;
                    return new OpenXmlValidator(version);
                };

                using (var presentation = PowerPointPresentation.Create(filePath)) {
                    Assert.Empty(presentation.DocumentValidationErrors);
                    Assert.Empty(presentation.DocumentValidationErrors);
                    Assert.Equal(1, factoryCalls);

                    presentation.Slides[0].AddTextBox("Cached validation");

                    Assert.Empty(presentation.DocumentValidationErrors);
                    Assert.Equal(2, factoryCalls);

                    Assert.Empty(presentation.ValidateDocument(forceRefresh: true));
                    Assert.Equal(3, factoryCalls);
                }
            } finally {
                PowerPointPresentation.ValidatorFactory = originalFactory;
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DocumentIsValid_UsesCachedValidationResults() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            int factoryCalls = 0;
            Func<FileFormatVersions, OpenXmlValidator> originalFactory = PowerPointPresentation.ValidatorFactory;

            try {
                PowerPointPresentation.ValidatorFactory = version => {
                    factoryCalls++;
                    return new OpenXmlValidator(version);
                };

                using (var presentation = PowerPointPresentation.Create(filePath)) {
                    Assert.True(presentation.DocumentIsValid);
                    Assert.True(presentation.DocumentIsValid);
                    Assert.Equal(1, factoryCalls);
                }
            } finally {
                PowerPointPresentation.ValidatorFactory = originalFactory;
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static string FormatValidationErrors(IEnumerable<ValidationErrorInfo> errors) {
            return string.Join(Environment.NewLine + Environment.NewLine,
                errors.Select(error =>
                    $"Description: {error.Description}\n" +
                    $"Id: {error.Id}\n" +
                    $"ErrorType: {error.ErrorType}\n" +
                    $"Part: {error.Part?.Uri}\n" +
                    $"Path: {error.Path?.XPath}"));
        }
    }
}

