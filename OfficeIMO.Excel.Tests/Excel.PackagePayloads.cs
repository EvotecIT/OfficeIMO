using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void PackagePayloads_VbaProjectCanBeInspectedExtractedAndRemoved() {
            string macroPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsm");
            string macroCopyPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsm");
            string rejectedCopyPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string macroFreePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            byte[] vbaProject = CreateVbaProjectPayload();

            try {
                using (ExcelDocument document = ExcelDocument.Create(macroPath)) {
                    document.AddWorksheet("Macro").CellValue(1, 1, "Enabled");
                    Assert.Equal(SpreadsheetDocumentType.MacroEnabledWorkbook, document.OpenXmlDocument.DocumentType);
                    document.AddMacro(vbaProject);

                    OfficeVbaProjectInfo info = document.InspectVbaProject(includeSha256: true)!;
                    Assert.Equal(vbaProject.Length, info.Length);
                    Assert.Equal(64, info.Sha256!.Length);
                    Assert.Equal(new[] { "Module1", "ThisWorkbook" }, info.ModuleNames);
                    Assert.Throws<InvalidDataException>(() =>
                        document.InspectVbaProject(includeSha256: false,
                            maxBytes: info.Length - 1));
                    Assert.Equal(vbaProject, document.ExtractMacros());
                    document.Save();
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(macroPath, false)) {
                    Assert.Equal(SpreadsheetDocumentType.MacroEnabledWorkbook, package.DocumentType);
                    Assert.NotNull(package.WorkbookPart!.VbaProjectPart);
                }

                using (ExcelDocument document = ExcelDocument.Load(macroPath)) {
                    Assert.True(document.HasMacros);
                    document.SaveCopy(macroCopyPath);
                    Assert.Throws<InvalidOperationException>(() => document.SaveCopy(rejectedCopyPath));
                    Assert.False(File.Exists(rejectedCopyPath));
                    Assert.Throws<InvalidOperationException>(() => document.Save(macroFreePath));
                    document.RemoveMacros();
                    Assert.False(document.HasMacros);
                    document.Save(macroFreePath);
                }

                using SpreadsheetDocument macroFree = SpreadsheetDocument.Open(macroFreePath, false);
                Assert.Equal(SpreadsheetDocumentType.Workbook, macroFree.DocumentType);
                Assert.Null(macroFree.WorkbookPart!.VbaProjectPart);
                using SpreadsheetDocument macroCopy = SpreadsheetDocument.Open(macroCopyPath, false);
                Assert.Equal(SpreadsheetDocumentType.MacroEnabledWorkbook, macroCopy.DocumentType);
                Assert.NotNull(macroCopy.WorkbookPart!.VbaProjectPart);
            } finally {
                TryDelete(macroPath);
                TryDelete(macroCopyPath);
                TryDelete(rejectedCopyPath);
                TryDelete(macroFreePath);
            }
        }

        [Theory]
        [InlineData(".xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook)]
        [InlineData(".xltx", SpreadsheetDocumentType.Template)]
        [InlineData(".xltm", SpreadsheetDocumentType.MacroEnabledTemplate)]
        public void PackagePayloads_CreateAndLoadPreserveWorkbookDocumentType(
            string extension,
            SpreadsheetDocumentType expectedType) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.AddWorksheet("Type").CellValue(1, 1, expectedType.ToString());
                    Assert.Equal(expectedType, document.OpenXmlDocument.DocumentType);
                    document.Save();
                    Assert.Equal(ExcelSavePackageWriter.SimplePackage, document.LastSaveDiagnostics.Writer);
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                    Assert.Equal(expectedType, package.DocumentType);
                }
                using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                    Assert.Equal(expectedType, document.OpenXmlDocument.DocumentType);
                }
            } finally {
                TryDelete(filePath);
            }
        }

        [Fact]
        public void PackagePayloads_SaveCopyNormalizesAddInWorkbookType() {
            string sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string addInCopyPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlam");

            try {
                using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                    document.AddWorksheet("AddIn").CellValue(1, 1, "SaveCopy");
                    document.SaveCopy(addInCopyPath);
                }

                using SpreadsheetDocument package = SpreadsheetDocument.Open(addInCopyPath, false);
                Assert.Equal(SpreadsheetDocumentType.AddIn, package.DocumentType);
            } finally {
                TryDelete(sourcePath);
                TryDelete(addInCopyPath);
            }
        }

        [Fact]
        public void PackagePayloads_EmbeddedPackageCanBeHashedReplacedAndRemoved() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            byte[] original = System.Text.Encoding.UTF8.GetBytes("original embedded package");
            byte[] replacement = System.Text.Encoding.UTF8.GetBytes("replacement embedded package");

            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.AddWorksheet("Payload").CellValue(1, 1, "Embedded");
                    document.Save();
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                    WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                    EmbeddedPackagePart embeddedPart = worksheetPart.AddEmbeddedPackagePart(EmbeddedPackagePartType.Xlsx);
                    using (var stream = new MemoryStream(original, writable: false)) {
                        embeddedPart.FeedData(stream);
                    }

                    string relationshipId = worksheetPart.GetIdOfPart(embeddedPart);
                    worksheetPart.Worksheet.Append(new OleObjects(
                        $"<x:oleObjects xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:oleObject progId=\"Package\" shapeId=\"1025\" r:id=\"{relationshipId}\" /></x:oleObjects>"));
                    worksheetPart.Worksheet.Save();
                }

                string payloadId;
                using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                    OfficeEmbeddedPayloadInfo info = Assert.Single(document.GetEmbeddedPayloads(includeSha256: true));
                    payloadId = info.Id;
                    Assert.Equal(OfficeEmbeddedPayloadKind.EmbeddedPackage, info.Kind);
                    Assert.Equal(original.Length, info.Length);
                    Assert.Equal(64, info.Sha256!.Length);
                    Assert.Equal(original, document.ExtractEmbeddedPayload(info.Id));
                    Assert.Throws<InvalidDataException>(() => document.ExtractEmbeddedPayload(info.Id, original.Length - 1));
                    document.ReplaceEmbeddedPayload(info.Id, replacement);
                    document.Save();
                }

                using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                    Assert.Equal(replacement, document.ExtractEmbeddedPayload(payloadId));
                    ExcelFeatureReport featureReport = document.InspectFeatures();
                    Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, Assert.Single(featureReport.FindFeatures("Embedded packages")).SupportLevel);
                    Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, Assert.Single(featureReport.FindFeatures("OLE objects")).SupportLevel);
                    Assert.True(document.RemoveEmbeddedPayload(payloadId));
                    Assert.False(document.RemoveEmbeddedPayload(payloadId));
                    Assert.Empty(document.GetEmbeddedPayloads());
                    document.Save();
                }

                using SpreadsheetDocument saved = SpreadsheetDocument.Open(filePath, false);
                WorksheetPart savedSheet = saved.WorkbookPart!.WorksheetParts.Single();
                Assert.Empty(savedSheet.EmbeddedPackageParts);
                Assert.Empty(savedSheet.Worksheet.Elements<OleObjects>());
            } finally {
                TryDelete(filePath);
            }
        }

        private static byte[] CreateVbaProjectPayload() {
            return OfficeCompoundFileWriter.Write(new[] {
                new OfficeCompoundStream("VBA/dir", Array.Empty<byte>()),
                new OfficeCompoundStream("VBA/_VBA_PROJECT", Array.Empty<byte>()),
                new OfficeCompoundStream("VBA/Module1", System.Text.Encoding.UTF8.GetBytes("Sub Main()\nEnd Sub")),
                new OfficeCompoundStream("VBA/ThisWorkbook", System.Text.Encoding.UTF8.GetBytes("Option Explicit")),
                new OfficeCompoundStream("PROJECT", Array.Empty<byte>())
            });
        }
    }
}
