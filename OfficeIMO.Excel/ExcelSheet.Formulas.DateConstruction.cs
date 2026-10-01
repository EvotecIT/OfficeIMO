namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateDateValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 3 || !TryResolveFormulaOrNumericArguments(tokens, out var numbers)) return false;
            if (numbers.Any(value => double.IsNaN(value) || double.IsInfinity(value)
                || value < int.MinValue || value > int.MaxValue)) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            int year = (int)numbers[0], month = (int)numbers[1], day = (int)numbers[2];
            if (year >= 0 && year <= 1899) year += 1900;
            if (year < 1900 || year > 9999 || month == int.MinValue) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            try {
                // Add the day in serial space so DATE(1900,2,29) and DATE(1900,3,0)
                // retain Excel's fictitious leap day, which DateTime cannot represent.
                DateTime monthStart = new DateTime(year, 1, 1).AddMonths(month - 1);
                double serial = ToExcelDateSerial(monthStart) + day - 1d;
                double maximum = ToExcelDateSerial(new DateTime(9999, 12, 31));
                result = serial < 0d || serial > maximum
                    ? FormulaArgumentValue.Error("#NUM!")
                    : new FormulaArgumentValue(serial, InvariantNumberText.Get(serial));
            } catch (ArgumentException) {
                result = FormulaArgumentValue.Error("#NUM!");
            } catch (OverflowException) {
                result = FormulaArgumentValue.Error("#NUM!");
            }
            return true;
        }
    }
}
