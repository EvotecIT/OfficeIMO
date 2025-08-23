using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Excel {
    internal static partial class FluentWorkbook {
        public static void Example_FluentWorkbook_AutoFilter(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Testing AutoFilter scenarios with fluent API");
            
            // Scenario 1: Table with AutoFilter (default)
            string filePath1 = Path.Combine(folderPath, "FluentWorkbook_TableWithAutoFilter.xlsx");
            using (ExcelDocument document1 = ExcelDocument.Create(filePath1)) {
                document1.Execution.Mode = ExecutionMode.Sequential;

                document1.AsFluent()
                    .Sheet("TableWithFilter", s => s
                        .HeaderRow("Name", "Score", "Department")
                        .Row(r => r.Values("Alice", 93, "Engineering"))
                        .Row(r => r.Values("Bob", 88, "Marketing"))
                        .Row(r => r.Values("Charlie", 76, "Sales"))
                        .Table(t => t.Add("A1:C4", true, "DataTable"))  // AutoFilter included by default
                    )
                    .End()
                    .Save(false);
                
                Console.WriteLine("[✓] Created table with AutoFilter");
            }
            
            // Scenario 2: Table without AutoFilter
            string filePath2 = Path.Combine(folderPath, "FluentWorkbook_TableNoAutoFilter.xlsx");
            using (ExcelDocument document2 = ExcelDocument.Create(filePath2)) {
                document2.Execution.Mode = ExecutionMode.Sequential;

                document2.AsFluent()
                    .Sheet("TableNoFilter", s => s
                        .HeaderRow("Name", "Score", "Department")
                        .Row(r => r.Values("Alice", 93, "Engineering"))
                        .Row(r => r.Values("Bob", 88, "Marketing"))
                        .Row(r => r.Values("Charlie", 76, "Sales"))
                        .Table(t => t
                            .Add("A1:C4", true, "DataTable", includeAutoFilter: false)  // Explicitly disable AutoFilter
                        )
                    )
                    .End()
                    .Save(false);
                
                Console.WriteLine("[✓] Created table without AutoFilter");
            }
            
            // Scenario 3: Worksheet-level AutoFilter (no table)
            string filePath3 = Path.Combine(folderPath, "FluentWorkbook_WorksheetAutoFilter.xlsx");
            using (ExcelDocument document3 = ExcelDocument.Create(filePath3)) {
                document3.Execution.Mode = ExecutionMode.Sequential;

                document3.AsFluent()
                    .Sheet("WorksheetFilter", s => s
                        .HeaderRow("Name", "Score", "Department")
                        .Row(r => r.Values("Alice", 93, "Engineering"))
                        .Row(r => r.Values("Bob", 88, "Marketing"))
                        .Row(r => r.Values("Charlie", 76, "Sales"))
                        .AutoFilter("A1:C4")  // Worksheet-level AutoFilter
                    )
                    .End()
                    .Save(false);
                
                Console.WriteLine("[✓] Created worksheet with AutoFilter (no table)");
            }
            
            // Scenario 4: Table first (no filter), then AutoFilter - should add filter to table
            string filePath4 = Path.Combine(folderPath, "FluentWorkbook_TableThenFilter.xlsx");
            using (ExcelDocument document4 = ExcelDocument.Create(filePath4)) {
                document4.Execution.Mode = ExecutionMode.Sequential;

                document4.AsFluent()
                    .Sheet("TableThenFilter", s => s
                        .HeaderRow("Name", "Score", "Department")
                        .Row(r => r.Values("Alice", 93, "Engineering"))
                        .Row(r => r.Values("Bob", 88, "Marketing"))
                        .Row(r => r.Values("Charlie", 76, "Sales"))
                        .Table(t => t
                            .Add("A1:C4", true, "DataTable", includeAutoFilter: false)  // Table without filter
                        )
                        .AutoFilter("A1:C4")  // Should detect table and add filter to it
                    )
                    .End()
                    .Save(false);
                
                Console.WriteLine("[✓] Table first, then AutoFilter - filter added to table");
            }
            
            // Scenario 5: AutoFilter first, then Table - should migrate filter to table
            string filePath5 = Path.Combine(folderPath, "FluentWorkbook_FilterThenTable.xlsx");
            using (ExcelDocument document5 = ExcelDocument.Create(filePath5)) {
                document5.Execution.Mode = ExecutionMode.Sequential;

                document5.AsFluent()
                    .Sheet("FilterThenTable", s => s
                        .HeaderRow("Name", "Score", "Department")
                        .Row(r => r.Values("Alice", 93, "Engineering"))
                        .Row(r => r.Values("Bob", 88, "Marketing"))
                        .Row(r => r.Values("Charlie", 76, "Sales"))
                        .AutoFilter("A1:C4")  // Add worksheet filter first
                        .Table(t => t
                            .Add("A1:C4", true, "DataTable")  // Table with default includeAutoFilter=true should migrate the filter
                        )
                    )
                    .End()
                    .Save(false);
                
                Console.WriteLine("[✓] AutoFilter first, then Table - filter migrated to table");
                
                if (openExcel) {
                    document5.Open();
                }
            }
        }
    }
}