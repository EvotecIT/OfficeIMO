using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsFormulaFunctionWriterMetadata {
        private static readonly Dictionary<string, ushort> VariableFunctionIds = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase) {
            ["COUNT"] = 0x0000,
            ["SUM"] = 0x0004,
            ["AVERAGE"] = 0x0005,
            ["MIN"] = 0x0006,
            ["MAX"] = 0x0007,
            ["NPV"] = 0x000b,
            ["STDEV"] = 0x000c,
            ["LOOKUP"] = 0x001c,
            ["INDEX"] = 0x001d,
            ["AND"] = 0x0024,
            ["OR"] = 0x0025,
            ["VAR"] = 0x002e,
            ["LINEST"] = 0x0031,
            ["TREND"] = 0x0032,
            ["LOGEST"] = 0x0033,
            ["GROWTH"] = 0x0034,
            ["PV"] = 0x0038,
            ["FV"] = 0x0039,
            ["NPER"] = 0x003a,
            ["PMT"] = 0x003b,
            ["RATE"] = 0x003c,
            ["IRR"] = 0x003e,
            ["MATCH"] = 0x0040,
            ["WEEKDAY"] = 0x0046,
            ["OFFSET"] = 0x004e,
            ["ABSREF"] = 0x004f,
            ["RELREF"] = 0x0050,
            ["ARGUMENT"] = 0x0051,
            ["SEARCH"] = 0x0052,
            ["CALLER"] = 0x0059,
            ["DEREF"] = 0x005a,
            ["WINDOWS"] = 0x005b,
            ["DOCUMENTS"] = 0x005d,
            ["ACTIVE.CELL"] = 0x005e,
            ["SELECTION"] = 0x005f,
            ["RESULT"] = 0x0060,
            ["LINKS"] = 0x0067,
            ["LOG"] = 0x006d,
            ["HLOOKUP"] = 0x0065,
            ["VLOOKUP"] = 0x0066,
            ["GET.FORMULA"] = 0x006a,
            ["GET.NAME"] = 0x006b,
            ["LEFT"] = 0x0073,
            ["RIGHT"] = 0x0074,
            ["SUBSTITUTE"] = 0x0078,
            ["NAMES"] = 0x007a,
            ["DIRECTORY"] = 0x007b,
            ["FIND"] = 0x007c,
            ["CELL"] = 0x007d,
            ["INDIRECT"] = 0x0094,
            ["DDB"] = 0x0090,
            ["GET.DEF"] = 0x0091,
            ["REFTEXT"] = 0x0092,
            ["TEXTREF"] = 0x0093,
            ["GET.CHART.ITEM"] = 0x00a0,
            ["IPMT"] = 0x00a7,
            ["PPMT"] = 0x00a8,
            ["COUNTA"] = 0x00a9,
            ["FILES"] = 0x00a6,
            ["GET.BAR"] = 0x00b6,
            ["PRODUCT"] = 0x00b7,
            ["GET.CELL"] = 0x00b9,
            ["GET.WORKSPACE"] = 0x00ba,
            ["GET.WINDOW"] = 0x00bb,
            ["GET.DOCUMENT"] = 0x00bc,
            ["GET.NOTE"] = 0x00bf,
            ["USDOLLAR"] = 0x00cc,
            ["FINDB"] = 0x00cd,
            ["SEARCHB"] = 0x00ce,
            ["LEFTB"] = 0x00d0,
            ["RIGHTB"] = 0x00d1,
            ["RANK"] = 0x00d8,
            ["ADDRESS"] = 0x00db,
            ["DAYS360"] = 0x00dc,
            ["VDB"] = 0x00de,
            ["MEDIAN"] = 0x00e3,
            ["SUMPRODUCT"] = 0x00e4,
            ["GET.LINK.INFO"] = 0x00f2,
            ["TEXT.BOX"] = 0x00f3,
            ["GET.OBJECT"] = 0x00f6,
            ["DB"] = 0x00f7,
            ["EVALUATE"] = 0x0101,
            ["GET.TOOLBAR"] = 0x0102,
            ["GET.TOOL"] = 0x0103,
            ["GET.WORKBOOK"] = 0x010c,
            ["AVEDEV"] = 0x010d,
            ["BETADIST"] = 0x010e,
            ["BETAINV"] = 0x0110,
            ["SUMXMY2"] = 0x012f,
            ["SUMX2MY2"] = 0x0130,
            ["SUMX2PY2"] = 0x0131,
            ["PROB"] = 0x013d,
            ["DEVSQ"] = 0x013e,
            ["GEOMEAN"] = 0x013f,
            ["HARMEAN"] = 0x0140,
            ["SUMSQ"] = 0x0141,
            ["KURT"] = 0x0142,
            ["SKEW"] = 0x0143,
            ["ZTEST"] = 0x0144,
            ["PERCENTRANK"] = 0x0149,
            ["MODE"] = 0x014a,
            ["CONCATENATE"] = 0x0150,
            ["SUBTOTAL"] = 0x0158,
            ["SUMIF"] = 0x0159,
            ["SCENARIO.GET"] = 0x015c,
            ["OPTIONS.LISTS.GET"] = 0x015d,
            ["ROMAN"] = 0x0162,
            ["OPEN.DIALOG"] = 0x0163,
            ["SAVE.DIALOG"] = 0x0164,
            ["VIEW.GET"] = 0x0165,
            ["GETPIVOTDATA"] = 0x0166,
            ["HYPERLINK"] = 0x0167,
            ["AVERAGEA"] = 0x0169,
            ["MAXA"] = 0x016a,
            ["MINA"] = 0x016b,
            ["STDEVPA"] = 0x016c,
            ["VARPA"] = 0x016d,
            ["STDEVA"] = 0x016e,
            ["VARA"] = 0x016f,
            ["RTD"] = 0x017b
        };

        private static bool TryGetVariableFunctionId(string functionName, out ushort functionId) {
            return VariableFunctionIds.TryGetValue(functionName, out functionId);
        }

        private static bool IsSupportedVariableFunctionArgumentCountCore(ushort functionId, int parameterCount) {
            switch (functionId) {
                case 0x0000:
                case 0x0004:
                case 0x0005:
                case 0x0006:
                case 0x0007:
                case 0x000c:
                case 0x002e:
                case 0x00a9:
                case 0x00b7:
                case 0x00e3:
                case 0x00e4:
                case 0x010d:
                case 0x013e:
                case 0x013f:
                case 0x0140:
                case 0x0141:
                case 0x0142:
                case 0x0143:
                case 0x014a:
                case 0x0150:
                case 0x0169:
                case 0x016a:
                case 0x016b:
                case 0x016c:
                case 0x016d:
                case 0x016e:
                case 0x016f:
                    return parameterCount >= 1 && parameterCount <= 30;
                case 0x000b:
                    return parameterCount >= 2 && parameterCount <= 30;
                case 0x001c:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x001d:
                    return parameterCount >= 2 && parameterCount <= 4;
                case 0x0024:
                case 0x0025:
                    return parameterCount >= 1 && parameterCount <= 30;
                case 0x0031:
                case 0x0032:
                case 0x0033:
                case 0x0034:
                    return parameterCount >= 1 && parameterCount <= 4;
                case 0x0038:
                case 0x0039:
                case 0x003a:
                case 0x003b:
                case 0x010e:
                case 0x0110:
                    return parameterCount >= 3 && parameterCount <= 5;
                case 0x003c:
                    return parameterCount >= 3 && parameterCount <= 6;
                case 0x003e:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0040:
                case 0x0052:
                case 0x007c:
                case 0x00cd:
                case 0x00ce:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x007d:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0046:
                case 0x006d:
                case 0x0094:
                case 0x0073:
                case 0x0074:
                case 0x00cc:
                case 0x00d0:
                case 0x00d1:
                case 0x0162:
                case 0x0167:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0166:
                    return parameterCount >= 2 && parameterCount <= 30;
                case 0x017b:
                    return parameterCount >= 3 && parameterCount <= 30;
                case 0x004e:
                    return parameterCount >= 3 && parameterCount <= 5;
                case 0x004f:
                case 0x0050:
                    return parameterCount == 2;
                case 0x0051:
                    return parameterCount >= 0 && parameterCount <= 3;
                case 0x0059:
                case 0x005e:
                case 0x005f:
                    return parameterCount == 0;
                case 0x005a:
                    return parameterCount == 1;
                case 0x005b:
                case 0x005d:
                    return parameterCount >= 0 && parameterCount <= 2;
                case 0x0060:
                    return parameterCount >= 0 && parameterCount <= 1;
                case 0x0064:
                    return parameterCount >= 2 && parameterCount <= 30;
                case 0x0065:
                case 0x0066:
                    return parameterCount == 3 || parameterCount == 4;
                case 0x0067:
                    return parameterCount >= 0 && parameterCount <= 2;
                case 0x006a:
                    return parameterCount == 1;
                case 0x006b:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x0078:
                case 0x013d:
                    return parameterCount == 3 || parameterCount == 4;
                case 0x007a:
                    return parameterCount >= 0 && parameterCount <= 3;
                case 0x007b:
                    return parameterCount >= 0 && parameterCount <= 1;
                case 0x00d8:
                case 0x0144:
                case 0x0149:
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
                case 0x00a6:
                    return parameterCount >= 0 && parameterCount <= 2;
                case 0x00b6:
                    return parameterCount >= 0 && parameterCount <= 4;
                case 0x00b9:
                case 0x00bb:
                case 0x00bc:
                case 0x010c:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x00ba:
                    return parameterCount == 1;
                case 0x00bf:
                    return parameterCount >= 0 && parameterCount <= 3;
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
                case 0x012f:
                case 0x0130:
                case 0x0131:
                    return parameterCount == 2;
                case 0x0158:
                    return parameterCount >= 2 && parameterCount <= 254;
                case 0x0159:
                    return parameterCount == 2 || parameterCount == 3;
                case 0x015c:
                    return parameterCount == 1 || parameterCount == 2;
                case 0x015d:
                    return parameterCount == 1;
                case 0x0163:
                    return parameterCount >= 0 && parameterCount <= 4;
                case 0x0164:
                    return parameterCount >= 0 && parameterCount <= 5;
                case 0x0165:
                    return parameterCount == 1 || parameterCount == 2;
                default:
                    return false;
            }
        }
    }
}
