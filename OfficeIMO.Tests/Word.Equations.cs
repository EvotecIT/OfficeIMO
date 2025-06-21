using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithEquation() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedWithEquation.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(1, document.Equations.Count);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithExponentEquation() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedWithExponentEquation.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(1, document.Equations.Count);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithIntegralEquation() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedWithIntegralEquation.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:int><m:intPr/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:int></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(1, document.Equations.Count);
            }
        }
    }
}
