using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Equations {
        internal static void Example_AddEquation(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with equation");
            string filePath = System.IO.Path.Combine(folderPath, "EquationDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(openWord);
            }
        }

        internal static void Example_AddEquationExponent(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with exponent equation");
            string filePath = System.IO.Path.Combine(folderPath, "EquationExponent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(openWord);
            }
        }

        internal static void Example_AddEquationIntegral(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with integral equation");
            string filePath = System.IO.Path.Combine(folderPath, "EquationIntegral.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:int><m:intPr/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:int></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(openWord);
            }
        }
    }
}
