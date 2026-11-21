using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_RemoveHiddenData_ScrubsWorkbookAndReopensValid() {
            string filePath = Path.Combine(_directoryWithFiles, "HiddenData.Workbook.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddSheet("Data");
                for (int row = 1; row <= 4; row++) data.Cell(row, 1).SetValue("Row " + row);
                data.Cell(1, 5).SetValue("Kept");
                data.SetComment(1, 1, "Check the totals.", "Private Author");
                data.AddThreadedComment("A2", "Is this final?", "Private Reviewer");
                data.SetRowHidden(3, true);
                data.SetColumnHidden(4, true);
                ExcelSheet secret = document.AddSheet("Secret");
                secret.Cell(1, 1).SetValue("Salary");
                secret.SetHidden(true);
                document.SetNamedRange("SecretSalary", "'Secret'!$A$1", save: false);
                document.Properties.Creator = "Private Author";
                document.Properties.LastModifiedBy = "Private Editor";
                document.ApplicationProperties.Company = "Private Company";
                document.SetCustomDocumentProperty("ClientName", "Private Client");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                OfficeHiddenDataReport found = document.InspectHiddenData();
                OfficeHiddenDataKind expected = OfficeHiddenDataKind.Comments | OfficeHiddenDataKind.PersonalIdentities |
                    OfficeHiddenDataKind.HiddenContent | OfficeHiddenDataKind.DocumentProperties | OfficeHiddenDataKind.CustomProperties;
                Assert.Equal(expected, found.Kinds & expected);
                Assert.Contains(found.GetFindings(OfficeHiddenDataKind.HiddenContent), finding => finding.Location == "Secret");
                Assert.Contains(found.GetFindings(OfficeHiddenDataKind.HiddenContent), finding => finding.Description.Contains("SecretSalary"));

                OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions { Kinds = OfficeHiddenDataKind.All });
                Assert.DoesNotContain(removed.Findings, finding => finding.Outcome == OfficeHiddenDataOutcome.Retained);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                Assert.Null(package.CustomFilePropertiesPart);
                Assert.Empty(package.WorkbookPart!.WorkbookPersonParts);
            }

            using (ExcelDocument reopened = ExcelDocument.Load(filePath)) {
                Assert.Empty(reopened.ValidateDocument());
                OfficeHiddenDataReport after = reopened.InspectHiddenData();
                Assert.False(after.HasHiddenData, string.Join(Environment.NewLine, after.Findings));
                ExcelSheet sheet = Assert.Single(reopened.Sheets);
                Assert.Empty(sheet.GetComments());
                Assert.Empty(sheet.GetThreadedComments());
                Assert.True(sheet.TryGetCellText(3, 1, out string shifted));
                Assert.Equal("Row 4", shifted);
                Assert.True(sheet.TryGetCellText(1, 4, out string kept));
                Assert.Equal("Kept", kept);
            }
        }

        [Fact]
        public void Test_InspectHiddenData_MatchesHiddenSheetNamesAsWholeReferences() {
            using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "HiddenData.SheetNames.xlsx"));
            document.AddSheet("RawData").Cell(1, 1).SetValue(1);
            ExcelSheet hidden = document.AddSheet("Data");
            hidden.Cell(1, 1).SetValue(2);
            hidden.SetHidden(true);
            document.SetNamedRange("VisibleTotal", "RawData!$A$1", save: false);
            document.SetNamedRange("HiddenTotal", "Data!$A$1", save: false);

            var names = document.InspectHiddenData().GetFindings(OfficeHiddenDataKind.HiddenContent).Select(finding => finding.Description).ToList();

            Assert.Contains(names, description => description.Contains("'HiddenTotal'"));
            Assert.DoesNotContain(names, description => description.Contains("'VisibleTotal'"));
        }

        [Fact]
        public void Test_RemoveHiddenData_KeepsThreadedCommentPlaceholderAuthors() {
            using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "HiddenData.ThreadedPlaceholder.xlsx"));
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.AddThreadedComment("A1", "Is this final?", "Private Reviewer");
            string id = sheet.GetThreadedComments().Single().Id!;
            // Excel writes the legacy note behind a threaded comment with a "tc={id}" author.
            sheet.SetComment(1, 1, "[Threaded comment] Is this final?", "tc=" + id);
            sheet.SetComment(2, 1, "Check the totals.", "Private Author");

            OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions { Kinds = OfficeHiddenDataKind.PersonalIdentities });

            Assert.DoesNotContain(removed.Findings, finding => finding.Value != null && finding.Value.StartsWith("tc=", StringComparison.Ordinal));
            var authors = sheet.GetComments().ToDictionary(note => note.CellReference, note => note.Author);
            Assert.Equal("tc=" + id, authors["A1"]);
            Assert.Equal("Author", authors["A2"]);
        }

        [Fact]
        public void Test_RemoveHiddenData_KeepsPowerQueryDataMashup() {
            using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "HiddenData.DataMashup.xlsx"));
            document.AddSheet("Data");
            WorkbookPart workbookPart = document.OpenXmlDocument.WorkbookPart!;
            CustomXmlPart AddCustomXml(string xml) {
                CustomXmlPart part = workbookPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
                return part;
            }
            CustomXmlPart mashup = AddCustomXml("<DataMashup xmlns=\"http://schemas.microsoft.com/DataMashup\">AAAAAA==</DataMashup>");
            AddCustomXml("<client xmlns=\"urn:example:crm\">Private Client</client>");
            string mashupUri = mashup.Uri.ToString();

            OfficeHiddenDataReport removed = document.RemoveHiddenData();

            OfficeHiddenDataFinding kept = Assert.Single(removed.GetFindings(OfficeHiddenDataKind.CustomXml), finding => finding.Location == mashupUri);
            Assert.Equal(OfficeHiddenDataOutcome.Retained, kept.Outcome);
            Assert.Equal(mashupUri, Assert.Single(workbookPart.CustomXmlParts).Uri.ToString());
        }
    }
}
