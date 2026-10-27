using System;
using System.IO;
using OfficeIMO.Excel;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Excel
{
    internal static partial class WorkbookAuthoring
    {
        public static void Example_Workbook_AutoFilter(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Testing AutoFilter scenarios with worksheet API");
            // Scenario 1: Table with AutoFilter (default)
            string filePath1 = Path.Combine(folderPath, "WorkbookAuthoring_TableWithAutoFilter.xlsx");
            using (ExcelDocument document1 = ExcelDocument.Create(filePath1))
            {
                document1.Execution.Mode = ExcelExecutionMode.Sequential;
                {
                    var sheet1 = document1.AddWorksheet("TableWithFilter");
                    int nextRow2 = 1;
                    sheet1.Cell(nextRow2, 1, "Name");
                    sheet1.Cell(nextRow2, 2, "Score");
                    sheet1.Cell(nextRow2, 3, "Department");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Alice");
                    sheet1.Cell(nextRow2, 2, 93);
                    sheet1.Cell(nextRow2, 3, "Engineering");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Bob");
                    sheet1.Cell(nextRow2, 2, 88);
                    sheet1.Cell(nextRow2, 3, "Marketing");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Charlie");
                    sheet1.Cell(nextRow2, 2, 76);
                    sheet1.Cell(nextRow2, 3, "Sales");
                    nextRow2++;
                    sheet1.AddTable("A1:C4", true, "DataTable");
                    document1.Save();
                }

                Console.WriteLine("[✓] Created table with AutoFilter");
            }

            // Scenario 2: Table without AutoFilter
            string filePath2 = Path.Combine(folderPath, "WorkbookAuthoring_TableNoAutoFilter.xlsx");
            using (ExcelDocument document2 = ExcelDocument.Create(filePath2))
            {
                document2.Execution.Mode = ExcelExecutionMode.Sequential;
                {
                    var sheet3 = document2.AddWorksheet("TableNoFilter");
                    int nextRow4 = 1;
                    sheet3.Cell(nextRow4, 1, "Name");
                    sheet3.Cell(nextRow4, 2, "Score");
                    sheet3.Cell(nextRow4, 3, "Department");
                    nextRow4++;
                    sheet3.Cell(nextRow4, 1, "Alice");
                    sheet3.Cell(nextRow4, 2, 93);
                    sheet3.Cell(nextRow4, 3, "Engineering");
                    nextRow4++;
                    sheet3.Cell(nextRow4, 1, "Bob");
                    sheet3.Cell(nextRow4, 2, 88);
                    sheet3.Cell(nextRow4, 3, "Marketing");
                    nextRow4++;
                    sheet3.Cell(nextRow4, 1, "Charlie");
                    sheet3.Cell(nextRow4, 2, 76);
                    sheet3.Cell(nextRow4, 3, "Sales");
                    nextRow4++;
                    sheet3.AddTable("A1:C4", true, "DataTable", includeAutoFilter: false);
                    document2.Save();
                }

                Console.WriteLine("[✓] Created table without AutoFilter");
            }

            // Scenario 3: Worksheet-level AutoFilter (no table)
            string filePath3 = Path.Combine(folderPath, "WorkbookAuthoring_WorksheetAutoFilter.xlsx");
            using (ExcelDocument document3 = ExcelDocument.Create(filePath3))
            {
                document3.Execution.Mode = ExcelExecutionMode.Sequential;
                {
                    var sheet5 = document3.AddWorksheet("WorksheetFilter");
                    int nextRow6 = 1;
                    sheet5.Cell(nextRow6, 1, "Name");
                    sheet5.Cell(nextRow6, 2, "Score");
                    sheet5.Cell(nextRow6, 3, "Department");
                    nextRow6++;
                    sheet5.Cell(nextRow6, 1, "Alice");
                    sheet5.Cell(nextRow6, 2, 93);
                    sheet5.Cell(nextRow6, 3, "Engineering");
                    nextRow6++;
                    sheet5.Cell(nextRow6, 1, "Bob");
                    sheet5.Cell(nextRow6, 2, 88);
                    sheet5.Cell(nextRow6, 3, "Marketing");
                    nextRow6++;
                    sheet5.Cell(nextRow6, 1, "Charlie");
                    sheet5.Cell(nextRow6, 2, 76);
                    sheet5.Cell(nextRow6, 3, "Sales");
                    nextRow6++;
                    sheet5.AddAutoFilter("A1:C4");
                    document3.Save();
                }

                Console.WriteLine("[✓] Created worksheet with AutoFilter (no table)");
            }

            // Scenario 4: Table first (no filter), then AutoFilter - should add filter to table
            string filePath4 = Path.Combine(folderPath, "WorkbookAuthoring_TableThenFilter.xlsx");
            using (ExcelDocument document4 = ExcelDocument.Create(filePath4))
            {
                document4.Execution.Mode = ExcelExecutionMode.Sequential;
                {
                    var sheet7 = document4.AddWorksheet("TableThenFilter");
                    int nextRow8 = 1;
                    sheet7.Cell(nextRow8, 1, "Name");
                    sheet7.Cell(nextRow8, 2, "Score");
                    sheet7.Cell(nextRow8, 3, "Department");
                    nextRow8++;
                    sheet7.Cell(nextRow8, 1, "Alice");
                    sheet7.Cell(nextRow8, 2, 93);
                    sheet7.Cell(nextRow8, 3, "Engineering");
                    nextRow8++;
                    sheet7.Cell(nextRow8, 1, "Bob");
                    sheet7.Cell(nextRow8, 2, 88);
                    sheet7.Cell(nextRow8, 3, "Marketing");
                    nextRow8++;
                    sheet7.Cell(nextRow8, 1, "Charlie");
                    sheet7.Cell(nextRow8, 2, 76);
                    sheet7.Cell(nextRow8, 3, "Sales");
                    nextRow8++;
                    sheet7.AddTable("A1:C4", true, "DataTable", includeAutoFilter: false);
                    sheet7.AddAutoFilter("A1:C4");
                    document4.Save();
                }

                Console.WriteLine("[✓] Table first, then AutoFilter - filter added to table");
            }

            // Scenario 5: AutoFilter first, then Table - should migrate filter to table
            string filePath5 = Path.Combine(folderPath, "WorkbookAuthoring_FilterThenTable.xlsx");
            using (ExcelDocument document5 = ExcelDocument.Create(filePath5))
            {
                document5.Execution.Mode = ExcelExecutionMode.Sequential;
                {
                    var sheet9 = document5.AddWorksheet("FilterThenTable");
                    int nextRow10 = 1;
                    sheet9.Cell(nextRow10, 1, "Name");
                    sheet9.Cell(nextRow10, 2, "Score");
                    sheet9.Cell(nextRow10, 3, "Department");
                    nextRow10++;
                    sheet9.Cell(nextRow10, 1, "Alice");
                    sheet9.Cell(nextRow10, 2, 93);
                    sheet9.Cell(nextRow10, 3, "Engineering");
                    nextRow10++;
                    sheet9.Cell(nextRow10, 1, "Bob");
                    sheet9.Cell(nextRow10, 2, 88);
                    sheet9.Cell(nextRow10, 3, "Marketing");
                    nextRow10++;
                    sheet9.Cell(nextRow10, 1, "Charlie");
                    sheet9.Cell(nextRow10, 2, 76);
                    sheet9.Cell(nextRow10, 3, "Sales");
                    nextRow10++;
                    sheet9.AddAutoFilter("A1:C4");
                    sheet9.AddTable("A1:C4", true, "DataTable");
                    document5.Save();
                }

                Console.WriteLine("[✓] AutoFilter first, then Table - filter migrated to table");
                if (openExcel)
                {
                    document5.OpenInApplication();
                }
            }
        }
    }
}