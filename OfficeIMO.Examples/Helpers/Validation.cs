using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Examples.Utils {
    internal static class Validation {
        internal static void ValidateDoc(string filePath) {
            try {
                using var doc = WordprocessingDocument.Open(filePath, false);
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(doc);
                bool ok = true;
                foreach (var err in errors) {
                    ok = false;
                    Console.WriteLine($"[!] Validation: {err.Description}\n    Part: {err.Part?.Uri}  Path: {err.Path?.XPath}\n    Id: {err.Id} ErrorType: {err.ErrorType}");
                }
                Console.WriteLine("[*] Document is valid: " + ok);
            } catch (Exception ex) {
                Console.WriteLine("[!] Validation exception: " + ex.Message);
            }
        }
    }
}
