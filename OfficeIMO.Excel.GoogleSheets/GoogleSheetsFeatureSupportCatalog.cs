namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>Support level for one Excel/Google Sheets translation capability.</summary>
    public enum GoogleSheetsFeatureSupportLevel {
        /// <summary>Represented as native editable content.</summary>
        Native = 0,
        /// <summary>Retained in a simpler target representation.</summary>
        Flattened = 1,
        /// <summary>Only the documented subset maps directly.</summary>
        Partial = 2,
        /// <summary>No supported direct translation is provided.</summary>
        Unsupported = 3,
    }

    /// <summary>Code-owned support-matrix row used by documentation and preflight tools.</summary>
    public sealed class GoogleSheetsFeatureSupport {
        /// <summary>Creates a feature row with separate export and import support levels.</summary>
        public GoogleSheetsFeatureSupport(string feature, GoogleSheetsFeatureSupportLevel export, GoogleSheetsFeatureSupportLevel import, string notes) {
            Feature = feature;
            Export = export;
            Import = import;
            Notes = notes;
        }
        /// <summary>Gets the feature group described by this row.</summary>
        public string Feature { get; }
        /// <summary>Gets the support level when exporting to Google Sheets.</summary>
        public GoogleSheetsFeatureSupportLevel Export { get; }
        /// <summary>Gets the support level when importing into OfficeIMO.</summary>
        public GoogleSheetsFeatureSupportLevel Import { get; }
        /// <summary>Gets the feature-specific behavior and limits.</summary>
        public string Notes { get; }
    }

    /// <summary>Authoritative feature matrix for the current package version.</summary>
    public static class GoogleSheetsFeatureSupportCatalog {
        private static readonly IReadOnlyList<GoogleSheetsFeatureSupport> FeaturesValue = new[] {
            new GoogleSheetsFeatureSupport("Sheets, values, formulas, merges and named ranges", GoogleSheetsFeatureSupportLevel.Native, GoogleSheetsFeatureSupportLevel.Native, "Native import supports ranges and field masks; Drive XLSX export is the broad fallback."),
            new GoogleSheetsFeatureSupport("Core cell and rich-text formatting", GoogleSheetsFeatureSupportLevel.Native, GoogleSheetsFeatureSupportLevel.Native, "Theme identities are resolved to colors when the source API exposes only rendered RGB."),
            new GoogleSheetsFeatureSupport("Notes", GoogleSheetsFeatureSupportLevel.Flattened, GoogleSheetsFeatureSupportLevel.Flattened, "Excel comments are intentionally flattened to Google cell notes."),
            new GoogleSheetsFeatureSupport("Filters, validations and native tables", GoogleSheetsFeatureSupportLevel.Native, GoogleSheetsFeatureSupportLevel.Partial, "Native import preserves supported basic filters, validation, and table metadata; Drive export covers the remainder."),
            new GoogleSheetsFeatureSupport("Conditional formatting", GoogleSheetsFeatureSupportLevel.Partial, GoogleSheetsFeatureSupportLevel.Partial, "Boolean/custom-formula rules are native; data bars, icon sets, and Excel-only rules are reported."),
            new GoogleSheetsFeatureSupport("Charts", GoogleSheetsFeatureSupportLevel.Partial, GoogleSheetsFeatureSupportLevel.Partial, "Column, bar, line, area, pie, doughnut, and scatter families are native; unsupported families fail or follow caller policy."),
            new GoogleSheetsFeatureSupport("Pivot tables", GoogleSheetsFeatureSupportLevel.Partial, GoogleSheetsFeatureSupportLevel.Partial, "Basic row, column, and aggregate fields are native; calculated fields, grouping, and advanced filters are reported."),
            new GoogleSheetsFeatureSupport("Protection", GoogleSheetsFeatureSupportLevel.Partial, GoogleSheetsFeatureSupportLevel.Partial, "Editors, domain editing, warning-only mode, and unprotected subranges are supported; Excel operation flags are diagnostic."),
            new GoogleSheetsFeatureSupport("Row and column outlines", GoogleSheetsFeatureSupportLevel.Native, GoogleSheetsFeatureSupportLevel.Partial, "Outline levels become nested Google dimension groups."),
            new GoogleSheetsFeatureSupport("Slicers and smart chips", GoogleSheetsFeatureSupportLevel.Unsupported, GoogleSheetsFeatureSupportLevel.Partial, "No semantic inference is performed; native imports report these objects."),
            new GoogleSheetsFeatureSupport("Embedded drawings and images", GoogleSheetsFeatureSupportLevel.Unsupported, GoogleSheetsFeatureSupportLevel.Partial, "IMAGE formulas and Drive links are not treated as embedded Excel drawings."),
            new GoogleSheetsFeatureSupport("Print layout", GoogleSheetsFeatureSupportLevel.Unsupported, GoogleSheetsFeatureSupportLevel.Unsupported, "Google Sheets has no equivalent print header/footer model."),
        };

        /// <summary>Gets the current feature groups and direction-specific support levels.</summary>
        public static IReadOnlyList<GoogleSheetsFeatureSupport> Features => FeaturesValue;
    }
}
