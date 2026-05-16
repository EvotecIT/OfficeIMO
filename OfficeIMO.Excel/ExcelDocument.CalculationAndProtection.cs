using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Formula calculation and cached-result policy used during save.
        /// </summary>
        public ExcelCalculationOptions Calculation { get; } = new ExcelCalculationOptions();

        /// <summary>
        /// Returns true when workbook-level structure or window protection is present.
        /// </summary>
        public bool IsWorkbookProtected {
            get {
                var protection = WorkbookRoot.GetFirstChild<WorkbookProtection>();
                return protection != null && ((protection.LockStructure?.Value ?? false) || (protection.LockWindows?.Value ?? false));
            }
        }

        /// <summary>
        /// Protects workbook structure/window metadata. This is not file encryption.
        /// </summary>
        public void ProtectWorkbook(ExcelWorkbookProtectionOptions? options = null) {
            var opts = options ?? new ExcelWorkbookProtectionOptions();
            var workbook = WorkbookRoot;
            var protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection == null) {
                protection = new WorkbookProtection();
                var workbookViews = workbook.GetFirstChild<BookViews>();
                if (workbookViews != null) {
                    workbook.InsertBefore(protection, workbookViews);
                } else if (workbook.GetFirstChild<Sheets>() is Sheets sheets) {
                    workbook.InsertBefore(protection, sheets);
                } else if (workbook.GetFirstChild<WorkbookProperties>() is WorkbookProperties workbookProperties) {
                    workbook.InsertAfter(protection, workbookProperties);
                } else if (workbook.GetFirstChild<FileSharing>() is FileSharing fileSharing) {
                    workbook.InsertAfter(protection, fileSharing);
                } else if (workbook.GetFirstChild<FileVersion>() is FileVersion fileVersion) {
                    workbook.InsertAfter(protection, fileVersion);
                } else {
                    workbook.InsertAt(protection, 0);
                }
            }

            protection.LockStructure = opts.ProtectStructure;
            protection.LockWindows = opts.ProtectWindows;
            string? hash = ExcelProtectionHash.ResolveLegacyHash(opts.Password, opts.LegacyPasswordHash);
            if (hash != null) {
                protection.WorkbookPassword = hash;
            } else {
                protection.WorkbookPassword = null;
                protection.RemoveAttribute("workbookPassword", string.Empty);
            }
            workbook.Save();
        }

        /// <summary>
        /// Removes workbook-level structure/window protection metadata.
        /// </summary>
        public void UnprotectWorkbook() {
            var workbook = WorkbookRoot;
            var protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection != null) {
                workbook.RemoveChild(protection);
                workbook.Save();
            }
        }

        /// <summary>
        /// Marks all formulas dirty so Excel-compatible applications recalculate them on open.
        /// </summary>
        public void InvalidateFormulas() {
            foreach (var sheet in Sheets) {
                sheet.InvalidateFormulas();
            }

            ConfigureFullCalculationOnOpen();
        }

        /// <summary>
        /// Removes cached values from all formula cells.
        /// </summary>
        public void ClearCachedFormulaResults() {
            foreach (var sheet in Sheets) {
                sheet.ClearCachedFormulaResults();
            }
        }

        /// <summary>
        /// Evaluates supported formulas and writes cached values.
        /// </summary>
        public int RecalculateSupportedFormulas() {
            int count = 0;
            foreach (var sheet in Sheets) {
                count += sheet.RecalculateSupportedFormulas();
            }

            return count;
        }

        /// <summary>
        /// Requests a full workbook recalculation when the file is opened.
        /// </summary>
        public void ConfigureFullCalculationOnOpen() {
            var workbook = WorkbookRoot;
            var properties = workbook.GetFirstChild<CalculationProperties>();
            if (properties == null) {
                properties = new CalculationProperties();
                workbook.Append(properties);
            }

            properties.ForceFullCalculation = true;
            properties.FullCalculationOnLoad = true;
            properties.CalculationMode = CalculateModeValues.Auto;
            workbook.Save();
        }

        internal void ApplyCalculationPolicyBeforeSave() {
            if (Calculation.EvaluateFormulasBeforeSave) {
                RecalculateSupportedFormulas();
            } else if (Calculation.ClearCachedFormulaResultsBeforeSave) {
                ClearCachedFormulaResults();
            }

            if (Calculation.MarkFormulasDirtyBeforeSave) {
                InvalidateFormulas();
            }

            if (Calculation.ForceFullCalculationOnOpen) {
                ConfigureFullCalculationOnOpen();
            }
        }
    }
}
