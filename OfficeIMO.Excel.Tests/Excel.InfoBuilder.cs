using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public class ExcelInfoBuilderTests
    {
        [Fact]
        public void InfoBuilder_SetsCoreAndAppProperties()
        {
            string filePath = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    doc.BuiltinDocumentProperties.Title = "Report Title";
                    doc.BuiltinDocumentProperties.Creator = "Alice";
                    doc.ApplicationProperties.Company = "Contoso";
                    doc.ApplicationProperties.Manager = "Bob";
                    doc.BuiltinDocumentProperties.LastModifiedBy = "Alice";
                    doc.Save();
                }
            }

            using (var doc = ExcelDocument.Load(filePath))
            {
                Assert.Equal("Report Title", doc.BuiltinDocumentProperties.Title);
                Assert.Equal("Alice", doc.BuiltinDocumentProperties.Creator);
                Assert.Equal("Alice", doc.BuiltinDocumentProperties.LastModifiedBy);
                Assert.Equal("Contoso", doc.ApplicationProperties.Company);
                Assert.Equal("Bob", doc.ApplicationProperties.Manager);
            }

            File.Delete(filePath);
        }
    }
}