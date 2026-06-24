namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Provides stable names for BIFF parsed-formula tokens.
    /// </summary>
    internal static class BiffFormulaTokenInfo {
        internal static string GetTokenClassName(byte token) {
            byte tokenClass = (byte)(token & 0x60);
            return tokenClass switch {
                0x20 => "Value",
                0x40 => "Reference",
                0x60 => "Array",
                _ => "Base"
            };
        }

        internal static string GetTokenName(byte token) {
            return token switch {
                0x01 => "PtgExp",
                0x02 => "PtgTbl",
                0x03 => "PtgAdd",
                0x04 => "PtgSub",
                0x05 => "PtgMul",
                0x06 => "PtgDiv",
                0x07 => "PtgPower",
                0x08 => "PtgConcat",
                0x09 => "PtgLt",
                0x0A => "PtgLe",
                0x0B => "PtgEq",
                0x0C => "PtgGe",
                0x0D => "PtgGt",
                0x0E => "PtgNe",
                0x0F => "PtgIsect",
                0x10 => "PtgUnion",
                0x11 => "PtgRange",
                0x12 => "PtgUplus",
                0x13 => "PtgUminus",
                0x14 => "PtgPercent",
                0x15 => "PtgParen",
                0x16 => "PtgMissArg",
                0x17 => "PtgStr",
                0x18 => "PtgExtended",
                0x19 => "PtgAttr",
                0x1C => "PtgErr",
                0x1D => "PtgBool",
                0x1E => "PtgInt",
                0x1F => "PtgNum",
                0x20 or 0x40 or 0x60 => "PtgArray",
                0x21 or 0x41 or 0x61 => "PtgFunc",
                0x22 or 0x42 or 0x62 => "PtgFuncVar",
                0x23 or 0x43 or 0x63 => "PtgName",
                0x24 or 0x44 or 0x64 => "PtgRef",
                0x25 or 0x45 or 0x65 => "PtgArea",
                0x26 or 0x46 or 0x66 => "PtgMemArea",
                0x27 or 0x47 or 0x67 => "PtgMemErr",
                0x28 or 0x48 or 0x68 => "PtgMemNoMem",
                0x29 or 0x49 or 0x69 => "PtgMemFunc",
                0x2A or 0x4A or 0x6A => "PtgRefErr",
                0x2B or 0x4B or 0x6B => "PtgAreaErr",
                0x2C or 0x4C or 0x6C => "PtgRefN",
                0x2D or 0x4D or 0x6D => "PtgAreaN",
                0x39 or 0x59 or 0x79 => "PtgNameX",
                0x3A or 0x5A or 0x7A => "PtgRef3d",
                0x3B or 0x5B or 0x7B => "PtgArea3d",
                0x3C or 0x5C or 0x7C => "PtgRefErr3d",
                0x3D or 0x5D or 0x7D => "PtgAreaErr3d",
                _ => $"FormulaToken0x{token:X2}"
            };
        }
    }
}
