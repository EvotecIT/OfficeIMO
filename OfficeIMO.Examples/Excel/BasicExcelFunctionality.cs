using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Excel {
    public class BasicExcelFunctionality {

        public static void BasicExcel_Example1(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating standard Excel Document 1");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Excel 1.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {

                ExcelSheet sheet1 = document.AddWorkSheet("Test");

                ExcelSheet sheet2 = document.AddWorkSheet("Test2");

                ExcelSheet sheet3 = document.AddWorkSheet("Test3");

                document.Save(openExcel);
            }
        }
        public static void BasicExcel_Example2(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating standard Excel Document 2");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Excel 2.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath, "WorkSheet1")) {

                document.Save(openExcel);
            }
        }

        public static void BasicExcel_Example3(bool openExcel) {
            Console.WriteLine("[*] Excel - Reading standard Excel Document 1");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
            string filePath = System.IO.Path.Combine(documentPaths, "BasicExcel.xlsx");
            using (ExcelDocument document = ExcelDocument.Load(filePath)) {

                Console.WriteLine("Sheets count:" + document.Sheets.Count);

                document.Save(openExcel);
            }
        }
    }
}
