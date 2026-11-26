using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal readonly partial struct ExcelFormulaCell {
        internal ExcelFormulaCell(Cell? cell) => _element = cell;

        public static implicit operator ExcelFormulaCell(Cell? cell) => new(cell);

        private Cell? Sdk => (Cell?)_element;

        private CellFormula? SdkFormula => _element == null ? null : ExcelWorksheetScan.FindFormula((Cell)_element);

        /// <summary>Whether the cell holds a formula element of its own.</summary>
        internal bool HasFormula => SdkFormula != null;

        /// <summary>The formula text as stored; empty for a shared formula's dependent cells.</summary>
        internal string? FormulaText => SdkFormula?.Text;

        /// <summary>How the formula is stored; <see cref="ExcelFormulaKind.Normal"/> when the cell has none.</summary>
        internal ExcelFormulaKind FormulaKind {
            get {
                CellFormulaValues? type = SdkFormula?.FormulaType?.Value;
                if (type == null) return ExcelFormulaKind.Normal;
                if (type == CellFormulaValues.Shared) return ExcelFormulaKind.Shared;
                if (type == CellFormulaValues.Array) return ExcelFormulaKind.Array;
                if (type == CellFormulaValues.DataTable) return ExcelFormulaKind.DataTable;
                return ExcelFormulaKind.Normal;
            }
        }

        /// <summary>The shared formula group the formula belongs to.</summary>
        internal uint? SharedIndex => SdkFormula?.SharedIndex?.Value;

        /// <summary>The range the formula covers (array and shared formula masters).</summary>
        internal string? FormulaReference => SdkFormula?.Reference?.Value;

        /// <summary>Whether the formula asks to be calculated on the next pass.</summary>
        internal bool? CalculateCell => SdkFormula?.CalculateCell?.Value;

        /// <summary>The A1 reference the cell records.</summary>
        internal string? Reference => Sdk?.CellReference?.Value;

        /// <summary>The stored value text, or null when the cell has no value element.</summary>
        internal string? ValueText => _element == null ? null : ExcelWorksheetScan.FindValue((Cell)_element)?.Text;

        /// <summary>Whether the cell has a value element.</summary>
        internal bool HasValue => _element != null && ExcelWorksheetScan.FindValue((Cell)_element) != null;

        /// <summary>The stored type of the value; null when the cell names none (a number).</summary>
        internal ExcelCellValueType? ValueType {
            get {
                CellValues? type = Sdk?.DataType?.Value;
                if (type == null) return null;
                if (type == CellValues.Number) return ExcelCellValueType.Number;
                if (type == CellValues.Boolean) return ExcelCellValueType.Boolean;
                if (type == CellValues.Date) return ExcelCellValueType.Date;
                if (type == CellValues.Error) return ExcelCellValueType.Error;
                if (type == CellValues.SharedString) return ExcelCellValueType.SharedString;
                if (type == CellValues.InlineString) return ExcelCellValueType.InlineString;
                return ExcelCellValueType.String;
            }
        }

        /// <summary>The value metadata index (rich values and dynamic array results).</summary>
        internal uint? ValueMetaIndex => Sdk?.ValueMetaIndex?.Value;

        /// <summary>The cell metadata index (dynamic array formulas).</summary>
        internal uint? CellMetaIndex => Sdk?.CellMetaIndex?.Value;

        /// <summary>The cell's style index.</summary>
        internal uint? StyleIndex => Sdk?.StyleIndex?.Value;

        internal static CellValues ToSdk(ExcelCellValueType type) => type switch {
            ExcelCellValueType.Number => CellValues.Number,
            ExcelCellValueType.Boolean => CellValues.Boolean,
            ExcelCellValueType.Date => CellValues.Date,
            ExcelCellValueType.Error => CellValues.Error,
            ExcelCellValueType.SharedString => CellValues.SharedString,
            ExcelCellValueType.InlineString => CellValues.InlineString,
            _ => CellValues.String,
        };

        /// <summary>Sets whether the formula is calculated on the next pass, leaving an equal value as it is.</summary>
        internal void SetCalculateCell(bool calculate) {
            CellFormula formula = SdkFormula!;
            if (formula.CalculateCell?.Value != calculate) formula.CalculateCell = calculate;
        }

        /// <summary>Whether the cell holds anything a spilled value would cover: a value, text, a formula or metadata.</summary>
        internal bool HasContent {
            get {
                Cell? cell = Sdk;
                return cell != null && (ExcelWorksheetScan.FindValue(cell) != null || cell.InlineString != null
                    || ExcelWorksheetScan.FindFormula(cell) != null || cell.CellMetaIndex != null || cell.ValueMetaIndex != null);
            }
        }

        /// <summary>Whether the cell holds an inline string.</summary>
        internal bool HasInlineString => Sdk?.InlineString != null;

        /// <summary>Removes the formula's calculate-on-load flag.</summary>
        internal void ClearCalculateCell() => SdkFormula!.CalculateCell = null;
    }
}
