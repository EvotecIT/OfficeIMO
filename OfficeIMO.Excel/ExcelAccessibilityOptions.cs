namespace OfficeIMO.Excel {
    /// <summary>Options controlling Excel workbook accessibility inspection.</summary>
    /// <remarks>
    /// The rules follow the Excel checks in Microsoft's Office Accessibility Checker: alternative text, table header rows,
    /// merged cells, default sheet names, hidden sheets, hyperlink text, text contrast and blank cells used for layout.
    /// </remarks>
    public sealed class ExcelAccessibilityOptions {
        private double _minimumTextContrastRatio = 4.5D;
        private double _minimumLargeTextContrastRatio = 3D;
        private int _maximumCellCount = 1_000_000;

        /// <summary>Selected policy profile. Strict turns structural warnings into errors; heuristic findings stay warnings.</summary>
        public OfficeAccessibilityPolicyProfile Profile { get; set; } = OfficeAccessibilityPolicyProfile.Default;
        /// <summary>Whether the workbook must define a document title. Defaults to false; the Strict profile turns it on.</summary>
        public bool RequireDocumentTitle { get; set; }
        /// <summary>Whether pictures, charts and shapes need alternative text unless marked decorative. Defaults to true.</summary>
        public bool RequireAlternativeText { get; set; } = true;
        /// <summary>Whether Excel tables need a header row. Defaults to true.</summary>
        public bool RequireTableHeaders { get; set; } = true;
        /// <summary>Whether merged cell ranges are reported. Defaults to true.</summary>
        public bool CheckMergedCells { get; set; } = true;
        /// <summary>Whether worksheets with content that keep a default name such as <c>Sheet1</c> are reported. Defaults to true.</summary>
        public bool CheckDefaultSheetNames { get; set; } = true;
        /// <summary>Whether hidden or very hidden worksheets that contain cell values are reported. Defaults to true.</summary>
        public bool CheckHiddenSheets { get; set; } = true;
        /// <summary>Whether cell hyperlinks need meaningful visible text. Defaults to true.</summary>
        public bool CheckMeaningfulLinks { get; set; } = true;
        /// <summary>Whether cell font colours are checked against the cell fill. Defaults to true.</summary>
        public bool CheckContrast { get; set; } = true;
        /// <summary>Whether blank rows and columns inside a sheet's data area are reported as likely spacing. Defaults to true.</summary>
        public bool CheckBlankCellLayout { get; set; } = true;

        /// <summary>Minimum contrast ratio for normal text. Defaults to the WCAG AA value 4.5.</summary>
        public double MinimumTextContrastRatio {
            get => _minimumTextContrastRatio;
            set => _minimumTextContrastRatio = ValidateContrast(value, nameof(MinimumTextContrastRatio));
        }

        /// <summary>Minimum contrast ratio for large text (18 pt, or 14 pt bold). Defaults to the WCAG AA value 3.</summary>
        public double MinimumLargeTextContrastRatio {
            get => _minimumLargeTextContrastRatio;
            set => _minimumLargeTextContrastRatio = ValidateContrast(value, nameof(MinimumLargeTextContrastRatio));
        }

        /// <summary>Maximum number of cells read per worksheet by the contrast, content and blank-layout checks. Defaults to 1,000,000.</summary>
        public int MaximumCellCount {
            get => _maximumCellCount;
            set {
                if (value <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumCellCount), "Value must be positive.");
                _maximumCellCount = value;
            }
        }

        /// <summary>Creates options initialized for a built-in profile.</summary>
        public static ExcelAccessibilityOptions ForProfile(OfficeAccessibilityPolicyProfile profile) =>
            new ExcelAccessibilityOptions {
                Profile = profile,
                RequireDocumentTitle = profile == OfficeAccessibilityPolicyProfile.Strict
            };

        internal OfficeAccessibilitySeverity Structural =>
            Profile == OfficeAccessibilityPolicyProfile.Strict ? OfficeAccessibilitySeverity.Error : OfficeAccessibilitySeverity.Warning;

        private static double ValidateContrast(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 1D || value > 21D) {
                throw new ArgumentOutOfRangeException(name, "Contrast ratio must be between 1 and 21.");
            }
            return value;
        }
    }
}
