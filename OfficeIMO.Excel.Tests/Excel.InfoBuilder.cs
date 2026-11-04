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
                    doc.Properties.Title = "Report Title";
                    doc.Properties.Creator = "Alice";
                    doc.ApplicationProperties.Company = "Contoso";
                    doc.ApplicationProperties.Manager = "Bob";
                    doc.Properties.LastModifiedBy = "Alice";
                    doc.Save();
                }
            }

            using (var doc = ExcelDocument.Load(filePath))
            {
                Assert.Equal("Report Title", doc.Properties.Title);
                Assert.Equal("Alice", doc.Properties.Creator);
                Assert.Equal("Alice", doc.Properties.LastModifiedBy);
                Assert.Equal("Contoso", doc.ApplicationProperties.Company);
                Assert.Equal("Bob", doc.ApplicationProperties.Manager);
            }

            File.Delete(filePath);
        }
    }
}