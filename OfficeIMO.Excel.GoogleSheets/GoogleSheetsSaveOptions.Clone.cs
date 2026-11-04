using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    public sealed partial class GoogleSheetsSaveOptions {
        /// <summary>Creates an independent copy of the mutable export settings, retaining the progress observer.</summary>
        public GoogleSheetsSaveOptions Clone() {
            if (Location == null || FidelityPolicy == null || UnsupportedFeatures == null || Replace == null ||
                Formulas == null || Execution == null || Spreadsheet == null || Protection == null || Identity == null)
                throw new ArgumentException("Google Sheets export option groups cannot be null.");
            var copy = new GoogleSheetsSaveOptions {
                Title = Title,
                Location = new GoogleDriveFileLocation {
                    DriveId = Location.DriveId, FolderId = Location.FolderId,
                    ExistingFileId = Location.ExistingFileId, SharedDriveAware = Location.SharedDriveAware
                },
                FidelityPolicy = new GoogleWorkspaceFidelityPolicy { PreflightMode = FidelityPolicy.PreflightMode },
                UnsupportedFeatures = new GoogleSheetsUnsupportedFeatureOptions {
                    Charts = UnsupportedFeatures.Charts, PivotTables = UnsupportedFeatures.PivotTables,
                    PrintLayout = UnsupportedFeatures.PrintLayout
                },
                Replace = new GoogleSheetsReplaceOptions {
                    ConflictMode = Replace.ConflictMode, ExpectedDriveVersion = Replace.ExpectedDriveVersion
                },
                Formulas = new GoogleSheetsFormulaOptions {
                    UnsupportedFormulaMode = Formulas.UnsupportedFormulaMode,
                    TreatUnknownFunctionsAsUnsupported = Formulas.TreatUnknownFunctionsAsUnsupported
                },
                Execution = new GoogleSheetsExecutionOptions {
                    UseValuesBatchUpdate = Execution.UseValuesBatchUpdate,
                    MaxValueRangesPerRequest = Execution.MaxValueRangesPerRequest,
                    MaxStructuralRequestsPerBatch = Execution.MaxStructuralRequestsPerBatch, Progress = Execution.Progress
                },
                Spreadsheet = new GoogleSheetsSpreadsheetOptions {
                    Locale = Spreadsheet.Locale, TimeZone = Spreadsheet.TimeZone,
                    RecalculationInterval = Spreadsheet.RecalculationInterval
                },
                Protection = new GoogleSheetsProtectionOptions {
                    WarningOnly = Protection.WarningOnly, DomainUsersCanEdit = Protection.DomainUsersCanEdit
                },
                Identity = new GoogleSheetsIdentityOptions {
                    WriteDeveloperMetadata = Identity.WriteDeveloperMetadata, SourceKey = Identity.SourceKey,
                    SourceValue = Identity.SourceValue, SchemaKey = Identity.SchemaKey, SchemaValue = Identity.SchemaValue
                }
            };
            foreach (string code in FidelityPolicy.AcceptedDiagnosticCodes) copy.FidelityPolicy.AcceptedDiagnosticCodes.Add(code);
            foreach (var mapping in Formulas.FunctionMappings) copy.Formulas.FunctionMappings.Add(mapping.Key, mapping.Value);
            foreach (string editor in Protection.EditorEmailAddresses) copy.Protection.EditorEmailAddresses.Add(editor);
            foreach (var ranges in Protection.UnprotectedRangesBySheet)
                copy.Protection.UnprotectedRangesBySheet.Add(ranges.Key, new List<string>(ranges.Value));
            return copy;
        }
    }
}
