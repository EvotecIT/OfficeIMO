using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
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
        public void Test_PowerPointValidationCacheInvalidation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (var presentation = PowerPointPresentation.Create(filePath)) {
                var initialErrors = presentation.DocumentValidationErrors;
                var cachedErrors = presentation.DocumentValidationErrors;

                Assert.Same(initialErrors, cachedErrors);

                presentation.AddSlide();

                var refreshedErrors = presentation.DocumentValidationErrors;

                Assert.NotSame(initialErrors, refreshedErrors);
                Assert.True(presentation.DocumentIsValid);
            }

            File.Delete(filePath);
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

