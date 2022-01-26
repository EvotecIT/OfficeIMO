using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    public class BasicExcelFunctionality {

        public static void BasicExcel_Example1(string filePath, bool openExcel) {
            
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {

                ExcelSheet sheet = document.AddWorkSheet("Test");


                document.Save(openExcel);
            }
        }
        public static void BasicExcel_Example2(string filePath, bool openExcel) {
            using (ExcelDocument document = ExcelDocument.Create(filePath, "WorkSheet1")) {

                document.Save(openExcel);
            }
        }
    }
}
