using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning-time options for Excel to Google Sheets export.
    /// </summary>
    public sealed class GoogleSheetsSaveOptions {
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        public string? Title { get; set; }
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
        public GoogleSheetsUnsupportedFeatureOptions UnsupportedFeatures { get; set; } = new GoogleSheetsUnsupportedFeatureOptions();
        public GoogleSheetsReplaceOptions Replace { get; set; } = new GoogleSheetsReplaceOptions();
        public GoogleSheetsFormulaOptions Formulas { get; set; } = new GoogleSheetsFormulaOptions();
        public GoogleSheetsExecutionOptions Execution { get; set; } = new GoogleSheetsExecutionOptions();
        public GoogleSheetsSpreadsheetOptions Spreadsheet { get; set; } = new GoogleSheetsSpreadsheetOptions();
        public GoogleSheetsProtectionOptions Protection { get; set; } = new GoogleSheetsProtectionOptions();
        public GoogleSheetsIdentityOptions Identity { get; set; } = new GoogleSheetsIdentityOptions();
    }

    public sealed class GoogleSheetsExecutionOptions {
        public bool UseValuesBatchUpdate { get; set; } = true;
        public int MaxValueRangesPerRequest { get; set; } = 100;
        public int MaxStructuralRequestsPerBatch { get; set; } = 400;
        public IProgress<GoogleSheetsExportProgress>? Progress { get; set; }
    }

    public sealed class GoogleSheetsExportProgress {
        public GoogleSheetsExportProgress(string stage, int completed, int total) {
            Stage = stage;
            Completed = completed;
            Total = total;
        }
        public string Stage { get; }
        public int Completed { get; }
        public int Total { get; }
    }

    public enum GoogleSheetsUnsupportedFormulaMode {
        Error = 0,
        PreserveWithWarning = 1,
        UseCachedValue = 2,
    }

    public sealed class GoogleSheetsFormulaOptions {
        public GoogleSheetsUnsupportedFormulaMode UnsupportedFormulaMode { get; set; } = GoogleSheetsUnsupportedFormulaMode.PreserveWithWarning;
        public bool TreatUnknownFunctionsAsUnsupported { get; set; } = true;
        public IDictionary<string, string> FunctionMappings { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class GoogleSheetsUnsupportedFeatureOptions {
        public UnsupportedFeatureMode Charts { get; set; } = UnsupportedFeatureMode.Error;
        public UnsupportedFeatureMode PivotTables { get; set; } = UnsupportedFeatureMode.Error;
        public UnsupportedFeatureMode PrintLayout { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
    }

    /// <summary>Google spreadsheet properties that have no reliable workbook equivalent.</summary>
    public sealed class GoogleSheetsSpreadsheetOptions {
        public string? Locale { get; set; }
        public string? TimeZone { get; set; }
        public GoogleSheetsRecalculationInterval RecalculationInterval { get; set; } = GoogleSheetsRecalculationInterval.OnChange;
    }

    public enum GoogleSheetsRecalculationInterval {
        OnChange = 0,
        Minute = 1,
        Hour = 2,
    }

    /// <summary>Target-side editors and unprotected subranges for translated Excel protection.</summary>
    public sealed class GoogleSheetsProtectionOptions {
        public bool WarningOnly { get; set; }
        public bool DomainUsersCanEdit { get; set; }
        public IList<string> EditorEmailAddresses { get; } = new List<string>();
        public IDictionary<string, IList<string>> UnprotectedRangesBySheet { get; } = new Dictionary<string, IList<string>>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Stable, discoverable metadata written into generated spreadsheets.</summary>
    public sealed class GoogleSheetsIdentityOptions {
        public bool WriteDeveloperMetadata { get; set; }
        public string SourceKey { get; set; } = "officeimo.source";
        public string SourceValue { get; set; } = "excel";
        public string SchemaKey { get; set; } = "officeimo.schema";
        public string SchemaValue { get; set; } = "1";
    }

    public enum GoogleSheetsReplaceConflictMode {
        RequireMatchingDriveVersion = 0,
        Overwrite = 1,
    }

    /// <summary>
    /// Safety contract for destructive replacement of an existing Google spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsReplaceOptions {
        public GoogleSheetsReplaceConflictMode ConflictMode { get; set; } = GoogleSheetsReplaceConflictMode.RequireMatchingDriveVersion;
        public long? ExpectedDriveVersion { get; set; }
    }
}
