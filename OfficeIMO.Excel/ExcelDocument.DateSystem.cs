using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Gets or sets the workbook date system used for serialized date values.
        /// </summary>
        public ExcelDateSystem DateSystem {
            get {
                WorkbookProperties? properties = WorkbookRoot.GetFirstChild<WorkbookProperties>();
                return properties?.Date1904?.Value == true
                    ? ExcelDateSystem.NineteenFour
                    : ExcelDateSystem.NineteenHundred;
            }
            set {
                if (value != ExcelDateSystem.NineteenHundred && value != ExcelDateSystem.NineteenFour) {
                    throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported Excel date system.");
                }

                Workbook workbook = WorkbookRoot;
                WorkbookProperties? properties = workbook.GetFirstChild<WorkbookProperties>();
                if (properties == null) {
                    properties = new WorkbookProperties();
                    OpenXmlWorkbookElementOrder.InsertInOrder(workbook, properties);
                }

                bool use1904 = value == ExcelDateSystem.NineteenFour;
                if (properties.Date1904?.Value == use1904) {
                    return;
                }

                properties.Date1904 = use1904 ? true : null;
                WorkbookRoot.Save();
                MarkPackageDirty();
            }
        }
    }
}
