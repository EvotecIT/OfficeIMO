using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning-time options for Excel to Google Sheets export.
    /// </summary>
    public sealed class GoogleSheetsSaveOptions {
        /// <summary>Gets or sets the Drive destination or existing spreadsheet target.</summary>
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        /// <summary>Gets or sets an optional spreadsheet title; a blank value uses the source workbook name or <c>Workbook</c>.</summary>
        public string? Title { get; set; }
        /// <summary>Gets or sets the policy checked against fidelity notices before mutation.</summary>
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
        /// <summary>Gets or sets handling for charts, pivots, and print layout without direct equivalents.</summary>
        public GoogleSheetsUnsupportedFeatureOptions UnsupportedFeatures { get; set; } = new GoogleSheetsUnsupportedFeatureOptions();
        /// <summary>Gets or sets the Drive-version preflight check used when replacing an existing spreadsheet.</summary>
        public GoogleSheetsReplaceOptions Replace { get; set; } = new GoogleSheetsReplaceOptions();
        /// <summary>Gets or sets formula translation and unsupported-function behavior.</summary>
        public GoogleSheetsFormulaOptions Formulas { get; set; } = new GoogleSheetsFormulaOptions();
        /// <summary>Gets or sets request batching and progress reporting choices.</summary>
        public GoogleSheetsExecutionOptions Execution { get; set; } = new GoogleSheetsExecutionOptions();
        /// <summary>Gets or sets target locale, time zone, and recalculation interval.</summary>
        public GoogleSheetsSpreadsheetOptions Spreadsheet { get; set; } = new GoogleSheetsSpreadsheetOptions();
        /// <summary>Gets or sets target-side protection editors and unprotected ranges.</summary>
        public GoogleSheetsProtectionOptions Protection { get; set; } = new GoogleSheetsProtectionOptions();
        /// <summary>Gets or sets optional developer metadata identifying generated spreadsheets.</summary>
        public GoogleSheetsIdentityOptions Identity { get; set; } = new GoogleSheetsIdentityOptions();
    }

    /// <summary>Controls how compiled requests are sent to Google Sheets.</summary>
    public sealed class GoogleSheetsExecutionOptions {
        /// <summary>Gets or sets whether cell values use the values batch-update endpoint; defaults to true.</summary>
        public bool UseValuesBatchUpdate { get; set; } = true;
        /// <summary>Gets or sets the maximum value ranges per request; defaults to 100.</summary>
        public int MaxValueRangesPerRequest { get; set; } = 100;
        /// <summary>Gets or sets the maximum structural requests per batch; defaults to 400.</summary>
        public int MaxStructuralRequestsPerBatch { get; set; } = 400;
        /// <summary>Gets or sets an optional observer for export stages.</summary>
        public IProgress<GoogleSheetsExportProgress>? Progress { get; set; }
    }

    /// <summary>Progress from one Google Sheets export stage.</summary>
    public sealed class GoogleSheetsExportProgress {
        /// <summary>Creates a progress snapshot with completed and total work counts.</summary>
        public GoogleSheetsExportProgress(string stage, int completed, int total) {
            Stage = stage;
            Completed = completed;
            Total = total;
        }
        /// <summary>Gets the export stage name.</summary>
        public string Stage { get; }
        /// <summary>Gets completed work for this stage.</summary>
        public int Completed { get; }
        /// <summary>Gets total work for this stage.</summary>
        public int Total { get; }
    }

    /// <summary>Action selected when a formula cannot be translated reliably.</summary>
    public enum GoogleSheetsUnsupportedFormulaMode {
        /// <summary>Report an error for unsupported formulas; whether it blocks export depends on the fidelity preflight mode and accepted diagnostic codes.</summary>
        Error = 0,
        /// <summary>Preserve the formula text and record a warning.</summary>
        PreserveWithWarning = 1,
        /// <summary>Write the cached value when one is available.</summary>
        UseCachedValue = 2,
    }

    /// <summary>Formula rewriting and unsupported-function policy.</summary>
    public sealed class GoogleSheetsFormulaOptions {
        /// <summary>Gets or sets how unsupported formulas are handled; defaults to preserving with a warning.</summary>
        public GoogleSheetsUnsupportedFormulaMode UnsupportedFormulaMode { get; set; } = GoogleSheetsUnsupportedFormulaMode.PreserveWithWarning;
        /// <summary>Gets or sets whether functions absent from the catalog count as unsupported; defaults to true.</summary>
        public bool TreatUnknownFunctionsAsUnsupported { get; set; } = true;
        /// <summary>Gets caller-supplied, case-insensitively keyed function-name replacements.</summary>
        public IDictionary<string, string> FunctionMappings { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Per-feature policy when no reliable native Google Sheets representation exists.</summary>
    public sealed class GoogleSheetsUnsupportedFeatureOptions {
        /// <summary>Gets or sets the unsupported-chart policy; defaults to error.</summary>
        public UnsupportedFeatureMode Charts { get; set; } = UnsupportedFeatureMode.Error;
        /// <summary>Gets or sets the unsupported-pivot policy; defaults to error.</summary>
        public UnsupportedFeatureMode PivotTables { get; set; } = UnsupportedFeatureMode.Error;
        /// <summary>Gets or sets the print-layout policy; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode PrintLayout { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
    }

    /// <summary>Google spreadsheet properties that have no reliable workbook equivalent.</summary>
    public sealed class GoogleSheetsSpreadsheetOptions {
        /// <summary>Gets or sets the target spreadsheet locale.</summary>
        public string? Locale { get; set; }
        /// <summary>Gets or sets the target spreadsheet time zone.</summary>
        public string? TimeZone { get; set; }
        /// <summary>Gets or sets the recalculation interval; defaults to on-change.</summary>
        public GoogleSheetsRecalculationInterval RecalculationInterval { get; set; } = GoogleSheetsRecalculationInterval.OnChange;
    }

    /// <summary>Recalculation frequency requested for the target spreadsheet.</summary>
    public enum GoogleSheetsRecalculationInterval {
        /// <summary>Recalculate when dependent values change.</summary>
        OnChange = 0,
        /// <summary>Recalculate approximately every minute.</summary>
        Minute = 1,
        /// <summary>Recalculate approximately every hour.</summary>
        Hour = 2,
    }

    /// <summary>Target-side editors and unprotected subranges for translated Excel protection.</summary>
    public sealed class GoogleSheetsProtectionOptions {
        /// <summary>Gets or sets whether protected ranges warn rather than block edits.</summary>
        public bool WarningOnly { get; set; }
        /// <summary>Gets or sets whether domain users may edit protected content.</summary>
        public bool DomainUsersCanEdit { get; set; }
        /// <summary>Gets the mutable list of editor email addresses for protected content.</summary>
        public IList<string> EditorEmailAddresses { get; } = new List<string>();
        /// <summary>Gets the mutable map of sheet names to unprotected A1 ranges.</summary>
        public IDictionary<string, IList<string>> UnprotectedRangesBySheet { get; } = new Dictionary<string, IList<string>>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Stable, discoverable metadata written into generated spreadsheets.</summary>
    public sealed class GoogleSheetsIdentityOptions {
        /// <summary>Gets or sets whether generated spreadsheets receive developer metadata; defaults to false.</summary>
        public bool WriteDeveloperMetadata { get; set; }
        /// <summary>Gets or sets the metadata key used for the source marker.</summary>
        public string SourceKey { get; set; } = "officeimo.source";
        /// <summary>Gets or sets the metadata value used for the source marker.</summary>
        public string SourceValue { get; set; } = "excel";
        /// <summary>Gets or sets the metadata key used for the schema marker.</summary>
        public string SchemaKey { get; set; } = "officeimo.schema";
        /// <summary>Gets or sets the metadata value used for the schema marker.</summary>
        public string SchemaValue { get; set; } = "1";
    }

    /// <summary>Drive-version behavior when replacing an existing spreadsheet.</summary>
    public enum GoogleSheetsReplaceConflictMode {
        /// <summary>Require the previously observed Drive version to match during preflight.</summary>
        RequireMatchingDriveVersion = 0,
        /// <summary>Replace without a matching-version guard.</summary>
        Overwrite = 1,
    }

    /// <summary>
    /// Safety contract for destructive replacement of an existing Google spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsReplaceOptions {
        /// <summary>Gets or sets the conflict behavior; defaults to checking a matching Drive version before mutation.</summary>
        /// <remarks>The exporter reads and compares the Drive version before writing; this is not an atomic conditional write.</remarks>
        public GoogleSheetsReplaceConflictMode ConflictMode { get; set; } = GoogleSheetsReplaceConflictMode.RequireMatchingDriveVersion;
        /// <summary>Gets or sets the Drive version observed before replacement.</summary>
        public long? ExpectedDriveVersion { get; set; }
    }
}
