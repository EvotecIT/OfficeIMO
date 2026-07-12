using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesSimpleExternalWorkbookLinkParts() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("Links");
                    sheet.CellValue(1, 1, "External link metadata");

                    ExternalWorkbookPart externalWorkbookPart = document.WorkbookPartRoot.AddNewPart<ExternalWorkbookPart>();
                    ExternalRelationship relationship = externalWorkbookPart.AddExternalRelationship(
                        ExternalLinkPathRelationshipType,
                        new Uri("Budget.xls", UriKind.Relative));
                    externalWorkbookPart.ExternalLink = new ExternalLink(new ExternalBook {
                        Id = relationship.Id,
                        SheetNames = new SheetNames(
                            new SheetName { Val = "Jan" },
                            new SheetName { Val = "Feb" }),
                        ExternalDefinedNames = new ExternalDefinedNames(
                            new ExternalDefinedName { Name = "TaxRate" },
                            new ExternalDefinedName { Name = "FebTaxRate", SheetId = 1U })
                    });
                    externalWorkbookPart.ExternalLink.Save();

                    document.WorkbookRoot.Append(new ExternalReferences(
                        new ExternalReference {
                            Id = document.WorkbookPartRoot.GetIdOfPart(externalWorkbookPart)
                        }));
                    document.WorkbookRoot.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                Assert.Equal(new[] { "Jan", "Feb" }, externalReference.SheetNames);

                Assert.Equal(2, externalReference.ExternalNames.Count);
                LegacyXlsExternalName workbookName = Assert.Single(externalReference.ExternalNames, name => name.Name == "TaxRate");
                Assert.Null(workbookName.LocalSheetIndex);
                Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, workbookName.BodyKind);

                LegacyXlsExternalName sheetName = Assert.Single(externalReference.ExternalNames, name => name.Name == "FebTaxRate");
                Assert.Equal(1, sheetName.LocalSheetIndex);
                Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, sheetName.BodyKind);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedExternalWorkbookLinkRelationshipTypesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("external workbook links", (document, sheet) => {
                sheet.CellValue(1, 1, "Unsupported external link relationship");

                ExternalWorkbookPart externalWorkbookPart = document.WorkbookPartRoot.AddNewPart<ExternalWorkbookPart>();
                ExternalRelationship relationship = externalWorkbookPart.AddExternalRelationship(
                    "http://example.com/notExternalLinkPath",
                    new Uri("Budget.xls", UriKind.Relative));
                externalWorkbookPart.ExternalLink = new ExternalLink(new ExternalBook {
                    Id = relationship.Id,
                    SheetNames = new SheetNames(new SheetName { Val = "Jan" })
                });
                externalWorkbookPart.ExternalLink.Save();

                document.WorkbookRoot.Append(new ExternalReferences(
                    new ExternalReference {
                        Id = document.WorkbookPartRoot.GetIdOfPart(externalWorkbookPart)
                    }));
                document.WorkbookRoot.Save();
            });
        }
    }
}
