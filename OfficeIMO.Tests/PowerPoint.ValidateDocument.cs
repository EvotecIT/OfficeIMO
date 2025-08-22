using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointValidateDocument {
        [Fact]
        public void Test_PowerPoint_ValidateDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                Assert.True(errors.Count > 0);
                Assert.False(presentation.DocumentIsValid);
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
