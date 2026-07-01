namespace OfficeIMO.Excel.LegacyXls.Write {
    /// <summary>
    /// Provides BIFF built-in function identifiers supported by the native XLS formula writer.
    /// </summary>
    internal static partial class LegacyXlsFormulaFunctionWriterMetadata {
        internal static bool TryGetVariableFunction(string formulaText, out ushort functionId, out int argumentStart) {
            functionId = 0;
            argumentStart = 0;
            if (!TryGetFunctionName(formulaText, out string? candidate, out argumentStart)) {
                return false;
            }

            return TryGetVariableFunctionId(candidate!, out functionId);
        }

        internal static bool IsSupportedVariableFunctionArgumentCount(ushort functionId, int parameterCount) {
            return IsSupportedVariableFunctionArgumentCountCore(functionId, parameterCount);
        }

        internal static bool AllowsMissingReferenceArgument(ushort functionId) {
            return functionId == 0x0008 || functionId == 0x0009;
        }

        internal static bool TryGetFixedFunction(string formulaText, out ushort functionId, out int parameterCount, out int argumentStart) {
            functionId = 0;
            parameterCount = 0;
            argumentStart = 0;
            if (!TryGetFunctionName(formulaText, out string? candidate, out argumentStart)) {
                return false;
            }

            switch (candidate!.ToUpperInvariant()) {
                case "ISNA":
                    functionId = 0x0002;
                    parameterCount = 1;
                    return true;
                case "ISERROR":
                    functionId = 0x0003;
                    parameterCount = 1;
                    return true;
                case "NA":
                    functionId = 0x000a;
                    parameterCount = 0;
                    return true;
                case "STDEV":
                    functionId = 0x000c;
                    parameterCount = 1;
                    return true;
                case "DOLLAR":
                    functionId = 0x000d;
                    parameterCount = 2;
                    return true;
                case "FIXED":
                    functionId = 0x000e;
                    parameterCount = 3;
                    return true;
                case "SIN":
                    functionId = 0x000f;
                    parameterCount = 1;
                    return true;
                case "COS":
                    functionId = 0x0010;
                    parameterCount = 1;
                    return true;
                case "TAN":
                    functionId = 0x0011;
                    parameterCount = 1;
                    return true;
                case "ATAN":
                    functionId = 0x0012;
                    parameterCount = 1;
                    return true;
                case "ROW":
                    functionId = 0x0008;
                    parameterCount = 1;
                    return true;
                case "COLUMN":
                    functionId = 0x0009;
                    parameterCount = 1;
                    return true;
                case "PI":
                    functionId = 0x0013;
                    parameterCount = 0;
                    return true;
                case "SQRT":
                    functionId = 0x0014;
                    parameterCount = 1;
                    return true;
                case "EXP":
                    functionId = 0x0015;
                    parameterCount = 1;
                    return true;
                case "LN":
                    functionId = 0x0016;
                    parameterCount = 1;
                    return true;
                case "LOG10":
                    functionId = 0x0017;
                    parameterCount = 1;
                    return true;
                case "ABS":
                    functionId = 0x0018;
                    parameterCount = 1;
                    return true;
                case "INT":
                    functionId = 0x0019;
                    parameterCount = 1;
                    return true;
                case "SIGN":
                    functionId = 0x001a;
                    parameterCount = 1;
                    return true;
                case "ROUND":
                    functionId = 0x001b;
                    parameterCount = 2;
                    return true;
                case "REPT":
                    functionId = 0x001e;
                    parameterCount = 2;
                    return true;
                case "MID":
                    functionId = 0x001f;
                    parameterCount = 3;
                    return true;
                case "LEN":
                    functionId = 0x0020;
                    parameterCount = 1;
                    return true;
                case "VALUE":
                    functionId = 0x0021;
                    parameterCount = 1;
                    return true;
                case "TRUE":
                    functionId = 0x0022;
                    parameterCount = 0;
                    return true;
                case "FALSE":
                    functionId = 0x0023;
                    parameterCount = 0;
                    return true;
                case "AND":
                    functionId = 0x0024;
                    parameterCount = 2;
                    return true;
                case "OR":
                    functionId = 0x0025;
                    parameterCount = 2;
                    return true;
                case "NOT":
                    functionId = 0x0026;
                    parameterCount = 1;
                    return true;
                case "MOD":
                    functionId = 0x0027;
                    parameterCount = 2;
                    return true;
                case "ATAN2":
                    functionId = 0x0061;
                    parameterCount = 2;
                    return true;
                case "ASIN":
                    functionId = 0x0062;
                    parameterCount = 1;
                    return true;
                case "ACOS":
                    functionId = 0x0063;
                    parameterCount = 1;
                    return true;
                case "DCOUNT":
                    functionId = 0x0028;
                    parameterCount = 3;
                    return true;
                case "DSUM":
                    functionId = 0x0029;
                    parameterCount = 3;
                    return true;
                case "DAVERAGE":
                    functionId = 0x002a;
                    parameterCount = 3;
                    return true;
                case "DMIN":
                    functionId = 0x002b;
                    parameterCount = 3;
                    return true;
                case "DMAX":
                    functionId = 0x002c;
                    parameterCount = 3;
                    return true;
                case "DSTDEV":
                    functionId = 0x002d;
                    parameterCount = 3;
                    return true;
                case "DVAR":
                    functionId = 0x002f;
                    parameterCount = 3;
                    return true;
                case "TEXT":
                    functionId = 0x0030;
                    parameterCount = 2;
                    return true;
                case "MIRR":
                    functionId = 0x003d;
                    parameterCount = 3;
                    return true;
                case "RAND":
                    functionId = 0x003f;
                    parameterCount = 0;
                    return true;
                case "DATE":
                    functionId = 0x0041;
                    parameterCount = 3;
                    return true;
                case "TIME":
                    functionId = 0x0042;
                    parameterCount = 3;
                    return true;
                case "DAY":
                    functionId = 0x0043;
                    parameterCount = 1;
                    return true;
                case "MONTH":
                    functionId = 0x0044;
                    parameterCount = 1;
                    return true;
                case "YEAR":
                    functionId = 0x0045;
                    parameterCount = 1;
                    return true;
                case "HOUR":
                    functionId = 0x0047;
                    parameterCount = 1;
                    return true;
                case "MINUTE":
                    functionId = 0x0048;
                    parameterCount = 1;
                    return true;
                case "SECOND":
                    functionId = 0x0049;
                    parameterCount = 1;
                    return true;
                case "NOW":
                    functionId = 0x004a;
                    parameterCount = 0;
                    return true;
                case "AREAS":
                    functionId = 0x004b;
                    parameterCount = 1;
                    return true;
                case "ROWS":
                    functionId = 0x004c;
                    parameterCount = 1;
                    return true;
                case "COLUMNS":
                    functionId = 0x004d;
                    parameterCount = 1;
                    return true;
                case "TRANSPOSE":
                    functionId = 0x0053;
                    parameterCount = 1;
                    return true;
                case "TYPE":
                    functionId = 0x0056;
                    parameterCount = 1;
                    return true;
                case "CHAR":
                    functionId = 0x006f;
                    parameterCount = 1;
                    return true;
                case "LOWER":
                    functionId = 0x0070;
                    parameterCount = 1;
                    return true;
                case "UPPER":
                    functionId = 0x0071;
                    parameterCount = 1;
                    return true;
                case "PROPER":
                    functionId = 0x0072;
                    parameterCount = 1;
                    return true;
                case "EXACT":
                    functionId = 0x0075;
                    parameterCount = 2;
                    return true;
                case "TRIM":
                    functionId = 0x0076;
                    parameterCount = 1;
                    return true;
                case "ISREF":
                    functionId = 0x0069;
                    parameterCount = 1;
                    return true;
                case "ISERR":
                    functionId = 0x007e;
                    parameterCount = 1;
                    return true;
                case "ISTEXT":
                    functionId = 0x007f;
                    parameterCount = 1;
                    return true;
                case "ISNUMBER":
                    functionId = 0x0080;
                    parameterCount = 1;
                    return true;
                case "ISBLANK":
                    functionId = 0x0081;
                    parameterCount = 1;
                    return true;
                case "T":
                    functionId = 0x0082;
                    parameterCount = 1;
                    return true;
                case "N":
                    functionId = 0x0083;
                    parameterCount = 1;
                    return true;
                case "REPLACE":
                    functionId = 0x0077;
                    parameterCount = 4;
                    return true;
                case "CODE":
                    functionId = 0x0079;
                    parameterCount = 1;
                    return true;
                case "REPLACEB":
                    functionId = 0x00cf;
                    parameterCount = 4;
                    return true;
                case "MIDB":
                    functionId = 0x00d2;
                    parameterCount = 3;
                    return true;
                case "LENB":
                    functionId = 0x00d3;
                    parameterCount = 1;
                    return true;
                case "ASC":
                    functionId = 0x00d6;
                    parameterCount = 1;
                    return true;
                case "DBCS":
                    functionId = 0x00d7;
                    parameterCount = 1;
                    return true;
                case "DATEVALUE":
                    functionId = 0x008c;
                    parameterCount = 1;
                    return true;
                case "TIMEVALUE":
                    functionId = 0x008d;
                    parameterCount = 1;
                    return true;
                case "SLN":
                    functionId = 0x008e;
                    parameterCount = 3;
                    return true;
                case "SYD":
                    functionId = 0x008f;
                    parameterCount = 4;
                    return true;
                case "CLEAN":
                    functionId = 0x00a2;
                    parameterCount = 1;
                    return true;
                case "MDETERM":
                    functionId = 0x00a3;
                    parameterCount = 1;
                    return true;
                case "MINVERSE":
                    functionId = 0x00a4;
                    parameterCount = 1;
                    return true;
                case "MMULT":
                    functionId = 0x00a5;
                    parameterCount = 2;
                    return true;
                case "FACT":
                    functionId = 0x00b8;
                    parameterCount = 1;
                    return true;
                case "DPRODUCT":
                    functionId = 0x00bd;
                    parameterCount = 3;
                    return true;
                case "DSTDEVP":
                    functionId = 0x00c3;
                    parameterCount = 3;
                    return true;
                case "DVARP":
                    functionId = 0x00c4;
                    parameterCount = 3;
                    return true;
                case "TRUNC":
                    functionId = 0x00c5;
                    parameterCount = 2;
                    return true;
                case "ROUNDUP":
                    functionId = 0x00d4;
                    parameterCount = 2;
                    return true;
                case "ROUNDDOWN":
                    functionId = 0x00d5;
                    parameterCount = 2;
                    return true;
                case "TODAY":
                    functionId = 0x00dd;
                    parameterCount = 0;
                    return true;
                case "DCOUNTA":
                    functionId = 0x00c7;
                    parameterCount = 3;
                    return true;
                case "DGET":
                    functionId = 0x00eb;
                    parameterCount = 3;
                    return true;
                case "INFO":
                    functionId = 0x00f4;
                    parameterCount = 1;
                    return true;
                case "FREQUENCY":
                    functionId = 0x00fc;
                    parameterCount = 2;
                    return true;
                case "SINH":
                    functionId = 0x00e5;
                    parameterCount = 1;
                    return true;
                case "COSH":
                    functionId = 0x00e6;
                    parameterCount = 1;
                    return true;
                case "TANH":
                    functionId = 0x00e7;
                    parameterCount = 1;
                    return true;
                case "ASINH":
                    functionId = 0x00e8;
                    parameterCount = 1;
                    return true;
                case "ACOSH":
                    functionId = 0x00e9;
                    parameterCount = 1;
                    return true;
                case "ATANH":
                    functionId = 0x00ea;
                    parameterCount = 1;
                    return true;
                case "STDEVP":
                    functionId = 0x00c1;
                    parameterCount = 1;
                    return true;
                case "VARP":
                    functionId = 0x00c2;
                    parameterCount = 1;
                    return true;
                case "CHITEST":
                    functionId = 0x0132;
                    parameterCount = 2;
                    return true;
                case "RSQ":
                    functionId = 0x0139;
                    parameterCount = 2;
                    return true;
                case "CORREL":
                    functionId = 0x0133;
                    parameterCount = 2;
                    return true;
                case "COVAR":
                    functionId = 0x0134;
                    parameterCount = 2;
                    return true;
                case "FORECAST":
                    functionId = 0x0135;
                    parameterCount = 3;
                    return true;
                case "FTEST":
                    functionId = 0x0136;
                    parameterCount = 2;
                    return true;
                case "INTERCEPT":
                    functionId = 0x0137;
                    parameterCount = 2;
                    return true;
                case "PEARSON":
                    functionId = 0x0138;
                    parameterCount = 2;
                    return true;
                case "STEYX":
                    functionId = 0x013a;
                    parameterCount = 2;
                    return true;
                case "SLOPE":
                    functionId = 0x013b;
                    parameterCount = 2;
                    return true;
                case "TTEST":
                    functionId = 0x013c;
                    parameterCount = 4;
                    return true;
                case "ISNONTEXT":
                    functionId = 0x00be;
                    parameterCount = 1;
                    return true;
                case "ISLOGICAL":
                    functionId = 0x00c6;
                    parameterCount = 1;
                    return true;
                case "LARGE":
                    functionId = 0x0145;
                    parameterCount = 2;
                    return true;
                case "SMALL":
                    functionId = 0x0146;
                    parameterCount = 2;
                    return true;
                case "QUARTILE":
                    functionId = 0x0147;
                    parameterCount = 2;
                    return true;
                case "PERCENTILE":
                    functionId = 0x0148;
                    parameterCount = 2;
                    return true;
                case "TRIMMEAN":
                    functionId = 0x014b;
                    parameterCount = 2;
                    return true;
                case "TINV":
                    functionId = 0x014c;
                    parameterCount = 2;
                    return true;
                case "POWER":
                    functionId = 0x0151;
                    parameterCount = 2;
                    return true;
                case "COUNTIF":
                    functionId = 0x015a;
                    parameterCount = 2;
                    return true;
                case "COUNTBLANK":
                    functionId = 0x015b;
                    parameterCount = 1;
                    return true;
                case "ERROR.TYPE":
                    functionId = 0x0105;
                    parameterCount = 1;
                    return true;
                case "ISPMT":
                    functionId = 0x015e;
                    parameterCount = 4;
                    return true;
                case "DATEDIF":
                    functionId = 0x015f;
                    parameterCount = 3;
                    return true;
                case "DATESTRING":
                    functionId = 0x0160;
                    parameterCount = 1;
                    return true;
                case "NUMBERSTRING":
                    functionId = 0x0161;
                    parameterCount = 2;
                    return true;
                case "PHONETIC":
                    functionId = 0x0168;
                    parameterCount = 1;
                    return true;
                case "BAHTTEXT":
                    functionId = 0x0170;
                    parameterCount = 1;
                    return true;
                case "THAIDAYOFWEEK":
                    functionId = 0x0171;
                    parameterCount = 1;
                    return true;
                case "THAIDIGIT":
                    functionId = 0x0172;
                    parameterCount = 1;
                    return true;
                case "THAIMONTHOFYEAR":
                    functionId = 0x0173;
                    parameterCount = 1;
                    return true;
                case "THAINUMSOUND":
                    functionId = 0x0174;
                    parameterCount = 1;
                    return true;
                case "THAINUMSTRING":
                    functionId = 0x0175;
                    parameterCount = 1;
                    return true;
                case "THAISTRINGLENGTH":
                    functionId = 0x0176;
                    parameterCount = 1;
                    return true;
                case "ISTHAIDIGIT":
                    functionId = 0x0177;
                    parameterCount = 1;
                    return true;
                case "ROUNDBAHTDOWN":
                    functionId = 0x0178;
                    parameterCount = 1;
                    return true;
                case "ROUNDBAHTUP":
                    functionId = 0x0179;
                    parameterCount = 1;
                    return true;
                case "THAIYEAR":
                    functionId = 0x017a;
                    parameterCount = 1;
                    return true;
                case "FLOOR":
                    functionId = 0x011d;
                    parameterCount = 2;
                    return true;
                case "CEILING":
                    functionId = 0x0120;
                    parameterCount = 2;
                    return true;
                case "GAMMALN":
                    functionId = 0x010f;
                    parameterCount = 1;
                    return true;
                case "BINOMDIST":
                    functionId = 0x0111;
                    parameterCount = 4;
                    return true;
                case "CHIDIST":
                    functionId = 0x0112;
                    parameterCount = 2;
                    return true;
                case "CHIINV":
                    functionId = 0x0113;
                    parameterCount = 2;
                    return true;
                case "COMBIN":
                    functionId = 0x0114;
                    parameterCount = 2;
                    return true;
                case "CONFIDENCE":
                    functionId = 0x0115;
                    parameterCount = 3;
                    return true;
                case "CRITBINOM":
                    functionId = 0x0116;
                    parameterCount = 3;
                    return true;
                case "EVEN":
                    functionId = 0x0117;
                    parameterCount = 1;
                    return true;
                case "EXPONDIST":
                    functionId = 0x0118;
                    parameterCount = 3;
                    return true;
                case "FDIST":
                    functionId = 0x0119;
                    parameterCount = 3;
                    return true;
                case "FINV":
                    functionId = 0x011a;
                    parameterCount = 3;
                    return true;
                case "FISHER":
                    functionId = 0x011b;
                    parameterCount = 1;
                    return true;
                case "FISHERINV":
                    functionId = 0x011c;
                    parameterCount = 1;
                    return true;
                case "GAMMADIST":
                    functionId = 0x011e;
                    parameterCount = 4;
                    return true;
                case "GAMMAINV":
                    functionId = 0x011f;
                    parameterCount = 3;
                    return true;
                case "HYPGEOMDIST":
                    functionId = 0x0121;
                    parameterCount = 4;
                    return true;
                case "LOGNORMDIST":
                    functionId = 0x0122;
                    parameterCount = 3;
                    return true;
                case "LOGINV":
                    functionId = 0x0123;
                    parameterCount = 3;
                    return true;
                case "NEGBINOMDIST":
                    functionId = 0x0124;
                    parameterCount = 3;
                    return true;
                case "NORMDIST":
                    functionId = 0x0125;
                    parameterCount = 4;
                    return true;
                case "NORMSDIST":
                    functionId = 0x0126;
                    parameterCount = 1;
                    return true;
                case "NORMINV":
                    functionId = 0x0127;
                    parameterCount = 3;
                    return true;
                case "NORMSINV":
                    functionId = 0x0128;
                    parameterCount = 1;
                    return true;
                case "STANDARDIZE":
                    functionId = 0x0129;
                    parameterCount = 3;
                    return true;
                case "ODD":
                    functionId = 0x012a;
                    parameterCount = 1;
                    return true;
                case "PERMUT":
                    functionId = 0x012b;
                    parameterCount = 2;
                    return true;
                case "POISSON":
                    functionId = 0x012c;
                    parameterCount = 3;
                    return true;
                case "TDIST":
                    functionId = 0x012d;
                    parameterCount = 3;
                    return true;
                case "WEIBULL":
                    functionId = 0x012e;
                    parameterCount = 4;
                    return true;
                case "RADIANS":
                    functionId = 0x0156;
                    parameterCount = 1;
                    return true;
                case "DEGREES":
                    functionId = 0x0157;
                    parameterCount = 1;
                    return true;
                default:
                    return false;
            }
        }

        internal static bool IsVolatileFixedFunction(ushort functionId) {
            return functionId == 0x003f
                || functionId == 0x004a
                || functionId == 0x00dd;
        }

        private static bool TryGetFunctionName(string formulaText, out string? functionName, out int argumentStart) {
            functionName = null;
            argumentStart = 0;
            int openParenthesis = formulaText.IndexOf('(');
            if (openParenthesis <= 0) {
                return false;
            }

            functionName = formulaText.Substring(0, openParenthesis).Trim();
            if (functionName.Length == 0) {
                return false;
            }

            argumentStart = openParenthesis + 1;
            return true;
        }
    }
}
