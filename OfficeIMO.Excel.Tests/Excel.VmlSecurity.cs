using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void HeaderFooterVmlRejectsDtdWhenReadingImages() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderFooterVmlDtd.xlsx");
            byte[] pngBytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sheet1");
                sheet.SetHeaderImage(HeaderFooterPosition.Center, pngBytes, "image/png");
                document.Save();
            }

            string relationshipId = GetSingleVmlImageRelationshipId(filePath);
            ReplaceFirstVmlPart(filePath, $$"""
                <?xml version="1.0" encoding="utf-8"?>
                <!DOCTYPE xml [<!ENTITY section "CH">]>
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <v:shape id="&section;" type="#_x0000_t75" style="width:96pt;height:32pt">
                    <v:imagedata r:id="{{relationshipId}}" o:relid="{{relationshipId}}" />
                  </v:shape>
                </xml>
                """);

            using ExcelDocument loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly });
            ExcelSheet.HeaderFooterSnapshot snapshot = loaded.Sheets.Single().GetHeaderFooter();

            Assert.True(snapshot.HeaderHasPicturePlaceholder);
            Assert.Null(snapshot.HeaderCenterImage);
        }

        [Fact]
        public void CommentVmlRejectsDtdBeforeUpdatingShapes() {
            string filePath = Path.Combine(_directoryWithFiles, "CommentVmlDtd.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sheet1");
                sheet.SetComment(1, 1, "Original");
                document.Save();
            }

            ReplaceFirstVmlPart(filePath, """
                <?xml version="1.0" encoding="utf-8"?>
                <!DOCTYPE xml [<!ENTITY shape "_x0000_s2048">]>
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                  <v:shape id="&shape;" type="#_x0000_t202">
                    <x:ClientData ObjectType="Note">
                      <x:Row>0</x:Row>
                      <x:Column>0</x:Column>
                    </x:ClientData>
                  </v:shape>
                </xml>
                """);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.Sheets.Single().SetComment(2, 1, "Updated");
                document.Save();
            }

            string vml = ReadFirstVmlPartText(filePath);
            Assert.DoesNotContain("<!DOCTYPE", vml, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("_x0000_s2049", vml, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("_x0000_s1025", vml, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetSingleVmlImageRelationshipId(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            VmlDrawingPart vmlPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts.Single();
            return vmlPart.Parts.Single(part => part.OpenXmlPart is ImagePart).RelationshipId;
        }

        private static void ReplaceFirstVmlPart(string filePath, string xml) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            VmlDrawingPart vmlPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts.Single();
            using Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
            using StreamWriter writer = new(stream, new UTF8Encoding(false));
            writer.Write(xml);
        }

        private static string ReadFirstVmlPartText(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            VmlDrawingPart vmlPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts.Single();
            using Stream stream = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
            using StreamReader reader = new(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }
    }
}
