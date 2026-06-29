namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents BIFF8 ignored formula error metadata parsed from an ISFFEC2 shared feature.
    /// </summary>
    public sealed class LegacyXlsIgnoredError {
        /// <summary>
        /// Creates ignored-error metadata for one or more worksheet ranges.
        /// </summary>
        public LegacyXlsIgnoredError(
            IReadOnlyList<string> references,
            bool evaluationError,
            bool emptyCellReference,
            bool numberStoredAsText,
            bool formulaRange,
            bool formula,
            bool twoDigitTextYear,
            bool unlockedFormula,
            bool listDataValidation) {
            References = references ?? throw new ArgumentNullException(nameof(references));
            EvaluationError = evaluationError;
            EmptyCellReference = emptyCellReference;
            NumberStoredAsText = numberStoredAsText;
            FormulaRange = formulaRange;
            Formula = formula;
            TwoDigitTextYear = twoDigitTextYear;
            UnlockedFormula = unlockedFormula;
            ListDataValidation = listDataValidation;
        }

        /// <summary>Gets the A1 references covered by this ignored-error rule.</summary>
        public IReadOnlyList<string> References { get; }

        /// <summary>Gets whether formula evaluation errors are ignored.</summary>
        public bool EvaluationError { get; }

        /// <summary>Gets whether references to empty cells are ignored.</summary>
        public bool EmptyCellReference { get; }

        /// <summary>Gets whether numbers stored as text are ignored.</summary>
        public bool NumberStoredAsText { get; }

        /// <summary>Gets whether formulas omitting part of a contiguous range are ignored.</summary>
        public bool FormulaRange { get; }

        /// <summary>Gets whether formulas inconsistent with neighboring formulas are ignored.</summary>
        public bool Formula { get; }

        /// <summary>Gets whether two-digit text year warnings are ignored.</summary>
        public bool TwoDigitTextYear { get; }

        /// <summary>Gets whether unlocked formula warnings are ignored.</summary>
        public bool UnlockedFormula { get; }

        /// <summary>Gets whether list data-validation warnings are ignored.</summary>
        public bool ListDataValidation { get; }
    }
}
