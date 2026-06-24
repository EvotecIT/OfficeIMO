namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Provides BIFF built-in function metadata used when projecting parsed formula tokens.
    /// </summary>
    internal static class BiffFormulaFunctionMetadata {
        internal static bool TryGetFixedFunctionMetadata(ushort functionId, out string? functionName, out int parameterCount) {
            switch (functionId) {
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
                case 0x0030:
                    functionName = "TEXT";
                    parameterCount = 2;
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
                case 0x004c:
                    functionName = "ROWS";
                    parameterCount = 1;
                    return true;
                case 0x004d:
                    functionName = "COLUMNS";
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
                case 0x00dd:
                    functionName = "TODAY";
                    parameterCount = 0;
                    return true;
                case 0x0139:
                    functionName = "RSQ";
                    parameterCount = 2;
                    return true;
                case 0x0145:
                    functionName = "LARGE";
                    parameterCount = 2;
                    return true;
                case 0x015a:
                    functionName = "COUNTIF";
                    parameterCount = 2;
                    return true;
                default:
                    functionName = null;
                    parameterCount = 0;
                    return false;
            }
        }

        internal static bool TryGetFunctionName(ushort functionId, out string? functionName) {
            switch (functionId) {
                case 0x0000:
                    functionName = "COUNT";
                    return true;
                case 0x0001:
                    functionName = "IF";
                    return true;
                case 0x0004:
                    functionName = "SUM";
                    return true;
                case 0x0005:
                    functionName = "AVERAGE";
                    return true;
                case 0x0006:
                    functionName = "MIN";
                    return true;
                case 0x0007:
                    functionName = "MAX";
                    return true;
                case 0x001d:
                    functionName = "INDEX";
                    return true;
                case 0x0040:
                    functionName = "MATCH";
                    return true;
                case 0x004e:
                    functionName = "OFFSET";
                    return true;
                case 0x0064:
                    functionName = "CHOOSE";
                    return true;
                case 0x0065:
                    functionName = "HLOOKUP";
                    return true;
                case 0x0066:
                    functionName = "VLOOKUP";
                    return true;
                case 0x0073:
                    functionName = "LEFT";
                    return true;
                case 0x0074:
                    functionName = "RIGHT";
                    return true;
                case 0x00a9:
                    functionName = "COUNTA";
                    return true;
                case 0x00b7:
                    functionName = "PRODUCT";
                    return true;
                case 0x00e3:
                    functionName = "MEDIAN";
                    return true;
                case 0x0158:
                    functionName = "SUBTOTAL";
                    return true;
                case 0x0159:
                    functionName = "SUMIF";
                    return true;
                default:
                    functionName = null;
                    return false;
            }
        }

        internal static bool TryGetKnownFunctionName(ushort functionId, out string? functionName) {
            if (TryGetFunctionName(functionId, out functionName)
                || TryGetFixedFunctionMetadata(functionId, out functionName, out _)) {
                return true;
            }

            switch (functionId) {
                case 0x000c:
                    functionName = "STDEV";
                    return true;
                default:
                    functionName = null;
                    return false;
            }
        }

        internal static bool IsSupportedVariableFunctionArgumentCount(ushort functionId, byte parameterCount) {
            switch (functionId) {
                case 0x0001:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x0065:
                case 0x0066:
                    return parameterCount == 3 || parameterCount == 4;
                case 0x0064:
                    return parameterCount >= 2 && parameterCount <= 30;
                case 0x001d:
                    return parameterCount == 2 || parameterCount == 3 || parameterCount == 4;
                case 0x0040:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x004e:
                    return parameterCount >= 3 && parameterCount <= 5;
                case 0x0158:
                    return parameterCount >= 2 && parameterCount <= 254;
                case 0x0159:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x0073:
                case 0x0074:
                    return parameterCount == 1 || parameterCount == 2;
                default:
                    return parameterCount > 0;
            }
        }
    }
}
