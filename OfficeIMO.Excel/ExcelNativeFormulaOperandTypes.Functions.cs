namespace OfficeIMO.Excel {
    internal static partial class ExcelNativeFormulaOperandTypes {
        private static readonly HashSet<string> ReferenceAggregates = new(StringComparer.OrdinalIgnoreCase) {
            "COUNT", "SUM", "AVERAGE", "MIN", "MAX", "STDEV", "VAR", "COUNTA", "PRODUCT", "STDEVP", "VARP", "MEDIAN",
            "AVEDEV", "DEVSQ", "GEOMEAN", "HARMEAN", "SUMSQ", "KURT", "SKEW", "MODE", "MAXA", "MINA", "STDEVPA", "VARPA", "STDEVA", "VARA", "AVERAGEA", "AND", "OR"
        };
        private static readonly HashSet<string> TwoArrays = new(StringComparer.OrdinalIgnoreCase) {
            "SUMPRODUCT", "SUMXMY2", "SUMX2MY2", "SUMX2PY2", "CORREL", "COVAR", "PEARSON", "RSQ", "SLOPE", "INTERCEPT", "STEYX", "FREQUENCY", "MMULT", "CHITEST", "FTEST", "TTEST"
        };
        private static readonly HashSet<string> FirstArray = new(StringComparer.OrdinalIgnoreCase) {
            "IRR", "MIRR", "LARGE", "SMALL", "QUARTILE", "PERCENTILE", "PERCENTRANK", "TRIMMEAN", "ZTEST", "MDETERM", "MINVERSE", "TRANSPOSE"
        };
        private static readonly HashSet<string> Databases = new(StringComparer.OrdinalIgnoreCase) {
            "DCOUNT", "DSUM", "DAVERAGE", "DMIN", "DMAX", "DSTDEV", "DVAR", "DCOUNTA", "DPRODUCT", "DSTDEVP", "DVARP", "DGET"
        };

        private static byte ArgumentType(string function, int index) {
            if (ReferenceAggregates.Contains(function)) return Reference;
            if (TwoArrays.Contains(function)) return function.Equals("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) || index < 2 ? Array : Value;
            if (FirstArray.Contains(function)) return index == 0 ? Array : Value;
            if (Databases.Contains(function)) return index is 0 or 2 ? Reference : Value;
            return function.ToUpperInvariant() switch {
                "IF" => index == 0 ? Value : (byte)0,
                "CHOOSE" => index == 0 ? Value : (byte)0,
                "NPV" or "SUBTOTAL" => index > 0 ? Reference : Value,
                "SUMIF" => index is 0 or 2 ? Reference : Value,
                "COUNTIF" or "COUNTBLANK" or "INDEX" or "OFFSET" or "ROW" or "COLUMN" or "ROWS" or "COLUMNS" or "AREAS" or "ISREF" => index == 0 ? Reference : Value,
                "LOOKUP" => index > 0 ? Array : Value,
                "VLOOKUP" or "HLOOKUP" or "MATCH" => index == 1 ? Array : Value,
                "RANK" or "CELL" or "GET.CELL" => index == 1 ? Reference : Value,
                "FORECAST" => index > 0 ? Array : Value,
                "LINEST" or "LOGEST" or "PROB" => index < 2 ? Array : Value,
                "TREND" or "GROWTH" => index < 3 ? Array : Value,
                _ => Value
            };
        }

        private static bool ReturnsReference(string function) => function is "INDEX" or "OFFSET" or "INDIRECT" or "IF" or "CHOOSE";
        private static bool ReturnsArray(string function) => function is "TRANSPOSE" or "MINVERSE" or "MMULT" or "FREQUENCY" or "LINEST" or "LOGEST" or "TREND" or "GROWTH";
    }
}
