using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
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

        public static void ValidateCorruptedWordDocument(string filepath) {
            // Insert some text into the body, this would cause Schema Error
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true)) {
                // Insert some text into the body, this would cause Schema Error
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                Run run = new Run(new Text("some text"));
                body.Append(run);

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
                    Console.WriteLine("count={0}", count);
                } catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
