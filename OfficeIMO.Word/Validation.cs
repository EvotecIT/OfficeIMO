using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class Validation {
        public static void ValidateWordDocument(string filepath) {
            Console.WriteLine("Validating document " + filepath);
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true)) {
                try {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    int count = 0;
                    foreach (ValidationErrorInfo error in validator.Validate(wordprocessingDocument)) {
                        count++;
                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }
                    Console.WriteLine("There were {0} errors found.", count);
                } catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
                wordprocessingDocument.Close();
            }
            Console.WriteLine("Validating document " + filepath);
        }
    }
}
