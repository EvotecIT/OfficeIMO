namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Provides BIFF built-in function metadata used when projecting parsed formula tokens.
    /// </summary>
    internal static partial class BiffFormulaFunctionMetadata {
        internal static bool TryGetFixedFunctionMetadata(ushort functionId, out string? functionName, out int parameterCount) {
            switch (functionId) {
                case 0x0002:
                    functionName = "ISNA";
                    parameterCount = 1;
                    return true;
                case 0x0003:
                    functionName = "ISERROR";
                    parameterCount = 1;
                    return true;
                case 0x0008:
                    functionName = "ROW";
                    parameterCount = 1;
                    return true;
                case 0x0009:
                    functionName = "COLUMN";
                    parameterCount = 1;
                    return true;
                case 0x000a:
                    functionName = "NA";
                    parameterCount = 0;
                    return true;
                case 0x000c:
                    functionName = "STDEV";
                    parameterCount = 1;
                    return true;
                case 0x000d:
                    functionName = "DOLLAR";
                    parameterCount = 2;
                    return true;
                case 0x000e:
                    functionName = "FIXED";
                    parameterCount = 3;
                    return true;
                case 0x000f:
                    functionName = "SIN";
                    parameterCount = 1;
                    return true;
                case 0x0010:
                    functionName = "COS";
                    parameterCount = 1;
                    return true;
                case 0x0011:
                    functionName = "TAN";
                    parameterCount = 1;
                    return true;
                case 0x0012:
                    functionName = "ATAN";
                    parameterCount = 1;
                    return true;
                case 0x0013:
                    functionName = "PI";
                    parameterCount = 0;
                    return true;
                case 0x0014:
                    functionName = "SQRT";
                    parameterCount = 1;
                    return true;
                case 0x0015:
                    functionName = "EXP";
                    parameterCount = 1;
                    return true;
                case 0x0016:
                    functionName = "LN";
                    parameterCount = 1;
                    return true;
                case 0x0017:
                    functionName = "LOG10";
                    parameterCount = 1;
                    return true;
                case 0x0018:
                    functionName = "ABS";
                    parameterCount = 1;
                    return true;
                case 0x0019:
                    functionName = "INT";
                    parameterCount = 1;
                    return true;
                case 0x001a:
                    functionName = "SIGN";
                    parameterCount = 1;
                    return true;
                case 0x001b:
                    functionName = "ROUND";
                    parameterCount = 2;
                    return true;
                case 0x001e:
                    functionName = "REPT";
                    parameterCount = 2;
                    return true;
                case 0x001f:
                    functionName = "MID";
                    parameterCount = 3;
                    return true;
                case 0x0020:
                    functionName = "LEN";
                    parameterCount = 1;
                    return true;
                case 0x0021:
                    functionName = "VALUE";
                    parameterCount = 1;
                    return true;
                case 0x0022:
                    functionName = "TRUE";
                    parameterCount = 0;
                    return true;
                case 0x0023:
                    functionName = "FALSE";
                    parameterCount = 0;
                    return true;
                case 0x0024:
                    functionName = "AND";
                    parameterCount = 2;
                    return true;
                case 0x0025:
                    functionName = "OR";
                    parameterCount = 2;
                    return true;
                case 0x0026:
                    functionName = "NOT";
                    parameterCount = 1;
                    return true;
                case 0x0027:
                    functionName = "MOD";
                    parameterCount = 2;
                    return true;
                case 0x0061:
                    functionName = "ATAN2";
                    parameterCount = 2;
                    return true;
                case 0x0062:
                    functionName = "ASIN";
                    parameterCount = 1;
                    return true;
                case 0x0063:
                    functionName = "ACOS";
                    parameterCount = 1;
                    return true;
                case 0x0028:
                    functionName = "DCOUNT";
                    parameterCount = 3;
                    return true;
                case 0x0029:
                    functionName = "DSUM";
                    parameterCount = 3;
                    return true;
                case 0x002a:
                    functionName = "DAVERAGE";
                    parameterCount = 3;
                    return true;
                case 0x002b:
                    functionName = "DMIN";
                    parameterCount = 3;
                    return true;
                case 0x002c:
                    functionName = "DMAX";
                    parameterCount = 3;
                    return true;
                case 0x002d:
                    functionName = "DSTDEV";
                    parameterCount = 3;
                    return true;
                case 0x002f:
                    functionName = "DVAR";
                    parameterCount = 3;
                    return true;
                case 0x0030:
                    functionName = "TEXT";
                    parameterCount = 2;
                    return true;
                case 0x003d:
                    functionName = "MIRR";
                    parameterCount = 3;
                    return true;
                case 0x003f:
                    functionName = "RAND";
                    parameterCount = 0;
                    return true;
                case 0x0041:
                    functionName = "DATE";
                    parameterCount = 3;
                    return true;
                case 0x0042:
                    functionName = "TIME";
                    parameterCount = 3;
                    return true;
                case 0x0043:
                    functionName = "DAY";
                    parameterCount = 1;
                    return true;
                case 0x0044:
                    functionName = "MONTH";
                    parameterCount = 1;
                    return true;
                case 0x0045:
                    functionName = "YEAR";
                    parameterCount = 1;
                    return true;
                case 0x0047:
                    functionName = "HOUR";
                    parameterCount = 1;
                    return true;
                case 0x0048:
                    functionName = "MINUTE";
                    parameterCount = 1;
                    return true;
                case 0x0049:
                    functionName = "SECOND";
                    parameterCount = 1;
                    return true;
                case 0x004a:
                    functionName = "NOW";
                    parameterCount = 0;
                    return true;
                case 0x004b:
                    functionName = "AREAS";
                    parameterCount = 1;
                    return true;
                case 0x004c:
                    functionName = "ROWS";
                    parameterCount = 1;
                    return true;
                case 0x004d:
                    functionName = "COLUMNS";
                    parameterCount = 1;
                    return true;
                case 0x0056:
                    functionName = "TYPE";
                    parameterCount = 1;
                    return true;
                case 0x0053:
                    functionName = "TRANSPOSE";
                    parameterCount = 1;
                    return true;
                case 0x006f:
                    functionName = "CHAR";
                    parameterCount = 1;
                    return true;
                case 0x0070:
                    functionName = "LOWER";
                    parameterCount = 1;
                    return true;
                case 0x0071:
                    functionName = "UPPER";
                    parameterCount = 1;
                    return true;
                case 0x0072:
                    functionName = "PROPER";
                    parameterCount = 1;
                    return true;
                case 0x0075:
                    functionName = "EXACT";
                    parameterCount = 2;
                    return true;
                case 0x0076:
                    functionName = "TRIM";
                    parameterCount = 1;
                    return true;
                case 0x0069:
                    functionName = "ISREF";
                    parameterCount = 1;
                    return true;
                case 0x0077:
                    functionName = "REPLACE";
                    parameterCount = 4;
                    return true;
                case 0x0079:
                    functionName = "CODE";
                    parameterCount = 1;
                    return true;
                case 0x00cf:
                    functionName = "REPLACEB";
                    parameterCount = 4;
                    return true;
                case 0x00d2:
                    functionName = "MIDB";
                    parameterCount = 3;
                    return true;
                case 0x00d3:
                    functionName = "LENB";
                    parameterCount = 1;
                    return true;
                case 0x00d6:
                    functionName = "ASC";
                    parameterCount = 1;
                    return true;
                case 0x00d7:
                    functionName = "DBCS";
                    parameterCount = 1;
                    return true;
                case 0x007e:
                    functionName = "ISERR";
                    parameterCount = 1;
                    return true;
                case 0x007f:
                    functionName = "ISTEXT";
                    parameterCount = 1;
                    return true;
                case 0x0080:
                    functionName = "ISNUMBER";
                    parameterCount = 1;
                    return true;
                case 0x0081:
                    functionName = "ISBLANK";
                    parameterCount = 1;
                    return true;
                case 0x0082:
                    functionName = "T";
                    parameterCount = 1;
                    return true;
                case 0x0083:
                    functionName = "N";
                    parameterCount = 1;
                    return true;
                case 0x008c:
                    functionName = "DATEVALUE";
                    parameterCount = 1;
                    return true;
                case 0x008d:
                    functionName = "TIMEVALUE";
                    parameterCount = 1;
                    return true;
                case 0x008e:
                    functionName = "SLN";
                    parameterCount = 3;
                    return true;
                case 0x008f:
                    functionName = "SYD";
                    parameterCount = 4;
                    return true;
                case 0x00a2:
                    functionName = "CLEAN";
                    parameterCount = 1;
                    return true;
                case 0x00a3:
                    functionName = "MDETERM";
                    parameterCount = 1;
                    return true;
                case 0x00a4:
                    functionName = "MINVERSE";
                    parameterCount = 1;
                    return true;
                case 0x00a5:
                    functionName = "MMULT";
                    parameterCount = 2;
                    return true;
                case 0x00b8:
                    functionName = "FACT";
                    parameterCount = 1;
                    return true;
                case 0x00bd:
                    functionName = "DPRODUCT";
                    parameterCount = 3;
                    return true;
                case 0x00c5:
                    functionName = "TRUNC";
                    parameterCount = 2;
                    return true;
                case 0x00d4:
                    functionName = "ROUNDUP";
                    parameterCount = 2;
                    return true;
                case 0x00d5:
                    functionName = "ROUNDDOWN";
                    parameterCount = 2;
                    return true;
                case 0x00dd:
                    functionName = "TODAY";
                    parameterCount = 0;
                    return true;
                case 0x00c1:
                    functionName = "STDEVP";
                    parameterCount = 1;
                    return true;
                case 0x00c2:
                    functionName = "VARP";
                    parameterCount = 1;
                    return true;
                case 0x00c3:
                    functionName = "DSTDEVP";
                    parameterCount = 3;
                    return true;
                case 0x00c4:
                    functionName = "DVARP";
                    parameterCount = 3;
                    return true;
                case 0x00be:
                    functionName = "ISNONTEXT";
                    parameterCount = 1;
                    return true;
                case 0x00c6:
                    functionName = "ISLOGICAL";
                    parameterCount = 1;
                    return true;
                case 0x00c7:
                    functionName = "DCOUNTA";
                    parameterCount = 3;
                    return true;
                case 0x00eb:
                    functionName = "DGET";
                    parameterCount = 3;
                    return true;
                case 0x00f4:
                    functionName = "INFO";
                    parameterCount = 1;
                    return true;
                case 0x00fc:
                    functionName = "FREQUENCY";
                    parameterCount = 2;
                    return true;
                case 0x00e5:
                    functionName = "SINH";
                    parameterCount = 1;
                    return true;
                case 0x00e6:
                    functionName = "COSH";
                    parameterCount = 1;
                    return true;
                case 0x00e7:
                    functionName = "TANH";
                    parameterCount = 1;
                    return true;
                case 0x00e8:
                    functionName = "ASINH";
                    parameterCount = 1;
                    return true;
                case 0x00e9:
                    functionName = "ACOSH";
                    parameterCount = 1;
                    return true;
                case 0x00ea:
                    functionName = "ATANH";
                    parameterCount = 1;
                    return true;
                case 0x011d:
                    functionName = "FLOOR";
                    parameterCount = 2;
                    return true;
                case 0x010f:
                    functionName = "GAMMALN";
                    parameterCount = 1;
                    return true;
                case 0x0111:
                    functionName = "BINOMDIST";
                    parameterCount = 4;
                    return true;
                case 0x0112:
                    functionName = "CHIDIST";
                    parameterCount = 2;
                    return true;
                case 0x0113:
                    functionName = "CHIINV";
                    parameterCount = 2;
                    return true;
                case 0x0114:
                    functionName = "COMBIN";
                    parameterCount = 2;
                    return true;
                case 0x0115:
                    functionName = "CONFIDENCE";
                    parameterCount = 3;
                    return true;
                case 0x0116:
                    functionName = "CRITBINOM";
                    parameterCount = 3;
                    return true;
                case 0x0117:
                    functionName = "EVEN";
                    parameterCount = 1;
                    return true;
                case 0x0118:
                    functionName = "EXPONDIST";
                    parameterCount = 3;
                    return true;
                case 0x0119:
                    functionName = "FDIST";
                    parameterCount = 3;
                    return true;
                case 0x011a:
                    functionName = "FINV";
                    parameterCount = 3;
                    return true;
                case 0x011b:
                    functionName = "FISHER";
                    parameterCount = 1;
                    return true;
                case 0x011c:
                    functionName = "FISHERINV";
                    parameterCount = 1;
                    return true;
                case 0x011e:
                    functionName = "GAMMADIST";
                    parameterCount = 4;
                    return true;
                case 0x011f:
                    functionName = "GAMMAINV";
                    parameterCount = 3;
                    return true;
                case 0x0121:
                    functionName = "HYPGEOMDIST";
                    parameterCount = 4;
                    return true;
                case 0x0122:
                    functionName = "LOGNORMDIST";
                    parameterCount = 3;
                    return true;
                case 0x0123:
                    functionName = "LOGINV";
                    parameterCount = 3;
                    return true;
                case 0x0124:
                    functionName = "NEGBINOMDIST";
                    parameterCount = 3;
                    return true;
                case 0x0125:
                    functionName = "NORMDIST";
                    parameterCount = 4;
                    return true;
                case 0x0126:
                    functionName = "NORMSDIST";
                    parameterCount = 1;
                    return true;
                case 0x0127:
                    functionName = "NORMINV";
                    parameterCount = 3;
                    return true;
                case 0x0128:
                    functionName = "NORMSINV";
                    parameterCount = 1;
                    return true;
                case 0x0129:
                    functionName = "STANDARDIZE";
                    parameterCount = 3;
                    return true;
                case 0x012a:
                    functionName = "ODD";
                    parameterCount = 1;
                    return true;
                case 0x012b:
                    functionName = "PERMUT";
                    parameterCount = 2;
                    return true;
                case 0x012c:
                    functionName = "POISSON";
                    parameterCount = 3;
                    return true;
                case 0x012d:
                    functionName = "TDIST";
                    parameterCount = 3;
                    return true;
                case 0x012e:
                    functionName = "WEIBULL";
                    parameterCount = 4;
                    return true;
                case 0x0120:
                    functionName = "CEILING";
                    parameterCount = 2;
                    return true;
                case 0x0132:
                    functionName = "CHITEST";
                    parameterCount = 2;
                    return true;
                case 0x0133:
                    functionName = "CORREL";
                    parameterCount = 2;
                    return true;
                case 0x0134:
                    functionName = "COVAR";
                    parameterCount = 2;
                    return true;
                case 0x0135:
                    functionName = "FORECAST";
                    parameterCount = 3;
                    return true;
                case 0x0136:
                    functionName = "FTEST";
                    parameterCount = 2;
                    return true;
                case 0x0137:
                    functionName = "INTERCEPT";
                    parameterCount = 2;
                    return true;
                case 0x0138:
                    functionName = "PEARSON";
                    parameterCount = 2;
                    return true;
                case 0x0139:
                    functionName = "RSQ";
                    parameterCount = 2;
                    return true;
                case 0x013a:
                    functionName = "STEYX";
                    parameterCount = 2;
                    return true;
                case 0x013b:
                    functionName = "SLOPE";
                    parameterCount = 2;
                    return true;
                case 0x013c:
                    functionName = "TTEST";
                    parameterCount = 4;
                    return true;
                case 0x0156:
                    functionName = "RADIANS";
                    parameterCount = 1;
                    return true;
                case 0x0157:
                    functionName = "DEGREES";
                    parameterCount = 1;
                    return true;
                case 0x0151:
                    functionName = "POWER";
                    parameterCount = 2;
                    return true;
                case 0x0145:
                    functionName = "LARGE";
                    parameterCount = 2;
                    return true;
                case 0x0146:
                    functionName = "SMALL";
                    parameterCount = 2;
                    return true;
                case 0x0147:
                    functionName = "QUARTILE";
                    parameterCount = 2;
                    return true;
                case 0x0148:
                    functionName = "PERCENTILE";
                    parameterCount = 2;
                    return true;
                case 0x014b:
                    functionName = "TRIMMEAN";
                    parameterCount = 2;
                    return true;
                case 0x014c:
                    functionName = "TINV";
                    parameterCount = 2;
                    return true;
                case 0x015a:
                    functionName = "COUNTIF";
                    parameterCount = 2;
                    return true;
                case 0x015b:
                    functionName = "COUNTBLANK";
                    parameterCount = 1;
                    return true;
                case 0x0105:
                    functionName = "ERROR.TYPE";
                    parameterCount = 1;
                    return true;
                case 0x015e:
                    functionName = "ISPMT";
                    parameterCount = 4;
                    return true;
                case 0x015f:
                    functionName = "DATEDIF";
                    parameterCount = 3;
                    return true;
                case 0x0160:
                    functionName = "DATESTRING";
                    parameterCount = 1;
                    return true;
                case 0x0161:
                    functionName = "NUMBERSTRING";
                    parameterCount = 2;
                    return true;
                case 0x0168:
                    functionName = "PHONETIC";
                    parameterCount = 1;
                    return true;
                case 0x0170:
                    functionName = "BAHTTEXT";
                    parameterCount = 1;
                    return true;
                case 0x0171:
                    functionName = "THAIDAYOFWEEK";
                    parameterCount = 1;
                    return true;
                case 0x0172:
                    functionName = "THAIDIGIT";
                    parameterCount = 1;
                    return true;
                case 0x0173:
                    functionName = "THAIMONTHOFYEAR";
                    parameterCount = 1;
                    return true;
                case 0x0174:
                    functionName = "THAINUMSOUND";
                    parameterCount = 1;
                    return true;
                case 0x0175:
                    functionName = "THAINUMSTRING";
                    parameterCount = 1;
                    return true;
                case 0x0176:
                    functionName = "THAISTRINGLENGTH";
                    parameterCount = 1;
                    return true;
                case 0x0177:
                    functionName = "ISTHAIDIGIT";
                    parameterCount = 1;
                    return true;
                case 0x0178:
                    functionName = "ROUNDBAHTDOWN";
                    parameterCount = 1;
                    return true;
                case 0x0179:
                    functionName = "ROUNDBAHTUP";
                    parameterCount = 1;
                    return true;
                case 0x017a:
                    functionName = "THAIYEAR";
                    parameterCount = 1;
                    return true;
                case 0x01e0:
                    functionName = "IFERROR";
                    parameterCount = 2;
                    return true;
                default:
                    functionName = null;
                    parameterCount = 0;
                    return false;
            }
        }

        internal static bool TryGetFunctionName(ushort functionId, out string? functionName) {
            return TryGetVariableFunctionName(functionId, out functionName);
        }

        internal static bool TryGetKnownFunctionName(ushort functionId, out string? functionName) {
            if (TryGetFunctionName(functionId, out functionName)
                || TryGetFixedFunctionMetadata(functionId, out functionName, out _)) {
                return true;
            }

            functionName = null;
            return false;
        }

        internal static bool IsSupportedVariableFunctionArgumentCount(ushort functionId, byte parameterCount) {
            return IsSupportedVariableFunctionArgumentCountCore(functionId, parameterCount);
        }
    }
}
