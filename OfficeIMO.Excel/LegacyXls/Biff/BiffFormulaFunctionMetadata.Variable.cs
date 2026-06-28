using System.Collections.Generic;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static partial class BiffFormulaFunctionMetadata {
        private static readonly Dictionary<ushort, string> VariableFunctionNames = new Dictionary<ushort, string> {
            [0x0000] = "COUNT",
            [0x0001] = "IF",
            [0x0004] = "SUM",
            [0x0005] = "AVERAGE",
            [0x0006] = "MIN",
            [0x0007] = "MAX",
            [0x000b] = "NPV",
            [0x000c] = "STDEV",
            [0x001c] = "LOOKUP",
            [0x001d] = "INDEX",
            [0x0024] = "AND",
            [0x0025] = "OR",
            [0x002e] = "VAR",
            [0x0031] = "LINEST",
            [0x0032] = "TREND",
            [0x0033] = "LOGEST",
            [0x0034] = "GROWTH",
            [0x0038] = "PV",
            [0x0039] = "FV",
            [0x003a] = "NPER",
            [0x003b] = "PMT",
            [0x003c] = "RATE",
            [0x003e] = "IRR",
            [0x0040] = "MATCH",
            [0x0046] = "WEEKDAY",
            [0x004e] = "OFFSET",
            [0x004f] = "ABSREF",
            [0x0050] = "RELREF",
            [0x0051] = "ARGUMENT",
            [0x0052] = "SEARCH",
            [0x0059] = "CALLER",
            [0x005a] = "DEREF",
            [0x005b] = "WINDOWS",
            [0x005d] = "DOCUMENTS",
            [0x005e] = "ACTIVE.CELL",
            [0x005f] = "SELECTION",
            [0x0060] = "RESULT",
            [0x0064] = "CHOOSE",
            [0x0065] = "HLOOKUP",
            [0x0066] = "VLOOKUP",
            [0x0067] = "LINKS",
            [0x006d] = "LOG",
            [0x006a] = "GET.FORMULA",
            [0x006b] = "GET.NAME",
            [0x0073] = "LEFT",
            [0x0074] = "RIGHT",
            [0x0078] = "SUBSTITUTE",
            [0x007a] = "NAMES",
            [0x007b] = "DIRECTORY",
            [0x007c] = "FIND",
            [0x007d] = "CELL",
            [0x0094] = "INDIRECT",
            [0x0090] = "DDB",
            [0x0091] = "GET.DEF",
            [0x0092] = "REFTEXT",
            [0x0093] = "TEXTREF",
            [0x00a0] = "GET.CHART.ITEM",
            [0x00a6] = "FILES",
            [0x00a9] = "COUNTA",
            [0x00a7] = "IPMT",
            [0x00a8] = "PPMT",
            [0x00b6] = "GET.BAR",
            [0x00b7] = "PRODUCT",
            [0x00b9] = "GET.CELL",
            [0x00ba] = "GET.WORKSPACE",
            [0x00bb] = "GET.WINDOW",
            [0x00bc] = "GET.DOCUMENT",
            [0x00bf] = "GET.NOTE",
            [0x00cc] = "USDOLLAR",
            [0x00cd] = "FINDB",
            [0x00ce] = "SEARCHB",
            [0x00d0] = "LEFTB",
            [0x00d1] = "RIGHTB",
            [0x00d8] = "RANK",
            [0x00db] = "ADDRESS",
            [0x00dc] = "DAYS360",
            [0x00de] = "VDB",
            [0x00e3] = "MEDIAN",
            [0x00e4] = "SUMPRODUCT",
            [0x00f2] = "GET.LINK.INFO",
            [0x00f3] = "TEXT.BOX",
            [0x00f6] = "GET.OBJECT",
            [0x00f7] = "DB",
            [0x0101] = "EVALUATE",
            [0x0102] = "GET.TOOLBAR",
            [0x0103] = "GET.TOOL",
            [0x010c] = "GET.WORKBOOK",
            [0x010d] = "AVEDEV",
            [0x010e] = "BETADIST",
            [0x0110] = "BETAINV",
            [0x012f] = "SUMXMY2",
            [0x0130] = "SUMX2MY2",
            [0x0131] = "SUMX2PY2",
            [0x013d] = "PROB",
            [0x013e] = "DEVSQ",
            [0x013f] = "GEOMEAN",
            [0x0140] = "HARMEAN",
            [0x0141] = "SUMSQ",
            [0x0142] = "KURT",
            [0x0143] = "SKEW",
            [0x0144] = "ZTEST",
            [0x0149] = "PERCENTRANK",
            [0x014a] = "MODE",
            [0x0150] = "CONCATENATE",
            [0x0158] = "SUBTOTAL",
            [0x0159] = "SUMIF",
            [0x015c] = "SCENARIO.GET",
            [0x015d] = "OPTIONS.LISTS.GET",
            [0x0162] = "ROMAN",
            [0x0163] = "OPEN.DIALOG",
            [0x0164] = "SAVE.DIALOG",
            [0x0165] = "VIEW.GET",
            [0x0166] = "GETPIVOTDATA",
            [0x0167] = "HYPERLINK",
            [0x0169] = "AVERAGEA",
            [0x016a] = "MAXA",
            [0x016b] = "MINA",
            [0x016c] = "STDEVPA",
            [0x016d] = "VARPA",
            [0x016e] = "STDEVA",
            [0x016f] = "VARA",
            [0x017b] = "RTD"
        };

        private static bool TryGetVariableFunctionName(ushort functionId, out string? functionName) {
            return VariableFunctionNames.TryGetValue(functionId, out functionName);
        }

        private static bool IsSupportedVariableFunctionArgumentCountCore(ushort functionId, byte parameterCount) {
            switch (functionId) {
                case 0x0001:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x000b:
                    return parameterCount >= 2;
                case 0x001c:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x0038:
                case 0x0039:
                case 0x003a:
                case 0x003b:
                    return parameterCount >= 3 && parameterCount <= 5;
                case 0x003c:
                    return parameterCount >= 3 && parameterCount <= 6;
                case 0x003e:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x010d:
                    return parameterCount >= 1 && parameterCount <= 30;
                case 0x010e:
                case 0x0110:
                    return parameterCount >= 3 && parameterCount <= 5;
                case 0x0065:
                case 0x0066:
                    return parameterCount == 3 || parameterCount == 4;
                case 0x0067:
                    return parameterCount <= 2;
                case 0x006a:
                    return parameterCount == 1;
                case 0x006b:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0064:
                    return parameterCount >= 2 && parameterCount <= 30;
                case 0x001d:
                    return parameterCount == 2 || parameterCount == 3 || parameterCount == 4;
                case 0x0024:
                case 0x0025:
                    return parameterCount >= 1 && parameterCount <= 30;
                case 0x0031:
                case 0x0032:
                case 0x0033:
                case 0x0034:
                    return parameterCount >= 1 && parameterCount <= 4;
                case 0x0040:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x004f:
                case 0x0050:
                    return parameterCount == 2;
                case 0x0051:
                    return parameterCount <= 3;
                case 0x0059:
                case 0x005e:
                case 0x005f:
                    return parameterCount == 0;
                case 0x005a:
                    return parameterCount == 1;
                case 0x005b:
                case 0x005d:
                    return parameterCount <= 2;
                case 0x0060:
                    return parameterCount <= 1;
                case 0x00d8:
                case 0x0144:
                case 0x0149:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x00dc:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x00db:
                    return parameterCount >= 2 && parameterCount <= 5;
                case 0x0090:
                case 0x00f7:
                    return parameterCount == 4 || parameterCount == 5;
                case 0x0091:
                    return parameterCount >= 1 && parameterCount <= 3;
                case 0x0092:
                case 0x0093:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x00a0:
                    return parameterCount >= 1 && parameterCount <= 3;
                case 0x00a7:
                case 0x00a8:
                    return parameterCount >= 4 && parameterCount <= 6;
                case 0x007a:
                    return parameterCount <= 3;
                case 0x007b:
                    return parameterCount <= 1;
                case 0x00a6:
                    return parameterCount <= 2;
                case 0x00b6:
                    return parameterCount <= 4;
                case 0x00b9:
                case 0x00bb:
                case 0x00bc:
                case 0x010c:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x00ba:
                    return parameterCount == 1;
                case 0x00bf:
                    return parameterCount <= 3;
                case 0x00de:
                    return parameterCount >= 5 && parameterCount <= 7;
                case 0x00f2:
                    return parameterCount >= 2 && parameterCount <= 4;
                case 0x00f3:
                    return parameterCount >= 1 && parameterCount <= 4;
                case 0x00f6:
                    return parameterCount >= 1 && parameterCount <= 5;
                case 0x0101:
                    return parameterCount == 1;
                case 0x0102:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0103:
                    return parameterCount >= 1 && parameterCount <= 3;
                case 0x013d:
                    return parameterCount == 3 || parameterCount == 4;
                case 0x0046:
                case 0x006d:
                case 0x0094:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x004e:
                    return parameterCount >= 3 && parameterCount <= 5;
                case 0x00cc:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0158:
                    return parameterCount >= 2 && parameterCount <= 254;
                case 0x0159:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x015c:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x015d:
                    return parameterCount == 1;
                case 0x0073:
                case 0x0074:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0052:
                case 0x007c:
                case 0x00cd:
                case 0x00ce:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x007d:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0078:
                    return parameterCount == 3 || parameterCount == 4;
                case 0x00d0:
                case 0x00d1:
                case 0x0162:
                case 0x0167:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0163:
                    return parameterCount <= 4;
                case 0x0164:
                    return parameterCount <= 5;
                case 0x0165:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0166:
                    return parameterCount >= 2 && parameterCount <= 30;
                case 0x017b:
                    return parameterCount >= 3 && parameterCount <= 30;
                case 0x0150:
                    return parameterCount >= 1;
                default:
                    return parameterCount > 0;
            }
        }
    }
}
