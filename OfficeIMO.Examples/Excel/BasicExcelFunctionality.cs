using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Contains basic examples demonstrating <see cref="ExcelDocument"/> usage.
    /// </summary>
    public class BasicExcelFunctionality {

        /// <summary>
        /// Creates a simple workbook with three sheets.
        /// </summary>
        /// <param name="folderPath">Target folder for the workbook.</param>
        /// <param name="openExcel">Opens the workbook after saving when set to <c>true</c>.</param>
        public static void BasicExcel_Example1(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating standard Excel Document 1");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Excel 1.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {

                ExcelSheet sheet1 = document.AddWorksheet("Test");

                ExcelSheet sheet2 = document.AddWorksheet("Test2");

                ExcelSheet sheet3 = document.AddWorksheet("Test3");

                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
        public static void BasicExcel_Example2(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating standard Excel Document 2");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Excel 2.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath, "Worksheet1")) {

                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }

        public static void BasicExcel_Example3(bool openExcel) {
            Console.WriteLine("[*] Excel - Reading standard Excel Document 1");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
            string filePath = System.IO.Path.Combine(documentPaths, "BasicExcel.xlsx");
            using (ExcelDocument document = ExcelDocument.Load(filePath)) {

                Console.WriteLine("Sheets count:" + document.Sheets.Count);

                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}
