namespace OfficeIMO.Excel.Xlsb.Biff12;

/// <summary>
/// Future functions (MS-XLSB 2.5.98.10 Ftab, Future Function table) that a formula calls through
/// PtgFuncVar 0x00FF with a PtgName to a BrtName whose fFutureFunction bit is set.
/// </summary>
internal static class XlsbFutureFunctions {
    /// <summary>The fixed BrtName flags of a future-function name: fHidden, fFunc, fProc and fFutureFunction.</summary>
    internal const uint NameFlags = 0x0002000BU;

    /// <summary>fFutureFunction in the BrtName flags.</summary>
    internal const uint FutureFunctionFlag = 0x00020000U;

    /// <summary>The text SpreadsheetML stores for a future-function defined name.</summary>
    internal const string NameFormulaText = "#NAME?";

    // LAMBDA, LET and the functions that only take LAMBDA arguments bind _xlpm. parameter names, and PY takes
    // Python-specific parameters; those need encodings beyond a plain call and are left out on purpose.
    private static readonly HashSet<string> Excluded = new(StringComparer.OrdinalIgnoreCase) {
        "_xlfn.LAMBDA", "_xlfn.LET", "_xlfn.MAKEARRAY", "_xlfn.MAP", "_xlfn.BYCOL", "_xlfn.BYROW", "_xlfn.ISOMITTED",
        "_xlfn._xlws.PY"
    };

    private static readonly Dictionary<string, string> Names = new[] {
        "_xlfn.ACOT", "_xlfn.ACOTH", "_xlfn.AGGREGATE", "_xlfn.ARABIC", "_xlfn.BASE", "_xlfn.BETA.DIST", "_xlfn.BETA.INV",
        "_xlfn.BINOM.DIST", "_xlfn.BINOM.DIST.RANGE", "_xlfn.BINOM.INV", "_xlfn.BITAND", "_xlfn.BITLSHIFT", "_xlfn.BITOR",
        "_xlfn.BITRSHIFT", "_xlfn.BITXOR", "_xlfn.BYCOL", "_xlfn.BYROW", "_xlfn.CEILING.MATH", "_xlfn.CEILING.PRECISE",
        "_xlfn.CHISQ.DIST", "_xlfn.CHISQ.DIST.RT", "_xlfn.CHISQ.INV", "_xlfn.CHISQ.INV.RT", "_xlfn.CHISQ.TEST",
        "_xlfn.CHOOSECOLS", "_xlfn.CHOOSEROWS", "_xlfn.COMBINA", "_xlfn.CONCAT", "_xlfn.CONFIDENCE.NORM",
        "_xlfn.CONFIDENCE.T", "_xlfn.COPILOT", "_xlfn.COT", "_xlfn.COTH", "_xlfn.COVARIANCE.P", "_xlfn.COVARIANCE.S",
        "_xlfn.CSC", "_xlfn.CSCH", "_xlfn.DAYS", "_xlfn.DECIMAL", "_xlfn.DROP", "_xlfn.ECMA.CEILING", "_xlfn.ERF.PRECISE",
        "_xlfn.ERFC.PRECISE", "_xlfn.EXPAND", "_xlfn.EXPON.DIST", "_xlfn.F.DIST", "_xlfn.F.DIST.RT", "_xlfn.F.INV",
        "_xlfn.F.INV.RT", "_xlfn.FIELDVALUE", "_xlfn.FILTERXML", "_xlfn.FLOOR.MATH", "_xlfn.FLOOR.PRECISE",
        "_xlfn.FORECAST.ETS", "_xlfn.FORECAST.ETS.CONFINT", "_xlfn.FORECAST.ETS.SEASONALITY", "_xlfn.FORECAST.ETS.STAT",
        "_xlfn.FORECAST.LINEAR", "_xlfn.FORMULATEXT", "_xlfn.GAMMA", "_xlfn.GAMMA.DIST", "_xlfn.GAMMA.INV",
        "_xlfn.GAMMALN.PRECISE", "_xlfn.GAUSS", "_xlfn.HSTACK", "_xlfn.HYPGEOM.DIST", "_xlfn.IFNA", "_xlfn.IFS",
        "_xlfn.IMCOSH", "_xlfn.IMCOT", "_xlfn.IMCSC", "_xlfn.IMCSCH", "_xlfn.IMSEC", "_xlfn.IMSECH", "_xlfn.IMSINH",
        "_xlfn.IMTAN", "_xlfn.ISFORMULA", "_xlfn.ISO.CEILING", "_xlfn.ISOMITTED", "_xlfn.ISOWEEKNUM", "_xlfn.LAMBDA",
        "_xlfn.LET", "_xlfn.LOGNORM.DIST", "_xlfn.LOGNORM.INV", "_xlfn.LONGTEXT", "_xlfn.MAKEARRAY", "_xlfn.MAP",
        "_xlfn.MAXIFS", "_xlfn.MINIFS", "_xlfn.MODE.MULT", "_xlfn.MODE.SNGL", "_xlfn.MUNIT", "_xlfn.NEGBINOM.DIST",
        "_xlfn.NETWORKDAYS.INTL", "_xlfn.NORM.DIST", "_xlfn.NORM.INV", "_xlfn.NORM.S.DIST", "_xlfn.NORM.S.INV",
        "_xlfn.NUMBERVALUE", "_xlfn.PDURATION", "_xlfn.PERCENTILE.EXC", "_xlfn.PERCENTILE.INC", "_xlfn.PERCENTRANK.EXC",
        "_xlfn.PERCENTRANK.INC", "_xlfn.PERMUTATIONA", "_xlfn.PHI", "_xlfn.POISSON.DIST", "_xlfn.PQSOURCE",
        "_xlfn.PYTHON_STR", "_xlfn.PYTHON_TYPE", "_xlfn.PYTHON_TYPENAME", "_xlfn.QUARTILE.EXC", "_xlfn.QUARTILE.INC",
        "_xlfn.QUERYSTRING", "_xlfn.RANDARRAY", "_xlfn.RANK.AVG", "_xlfn.RANK.EQ", "_xlfn.RRI", "_xlfn.SEC", "_xlfn.SECH",
        "_xlfn.SEQUENCE", "_xlfn.SHEET", "_xlfn.SHEETS", "_xlfn.SKEW.P", "_xlfn.SORTBY", "_xlfn.STDEV.P", "_xlfn.STDEV.S",
        "_xlfn.SWITCH", "_xlfn.T.DIST", "_xlfn.T.DIST.2T", "_xlfn.T.DIST.RT", "_xlfn.T.INV", "_xlfn.T.INV.2T",
        "_xlfn.T.TEST", "_xlfn.TAKE", "_xlfn.TEXTAFTER", "_xlfn.TEXTBEFORE", "_xlfn.TEXTJOIN", "_xlfn.TEXTSPLIT",
        "_xlfn.TOCOL", "_xlfn.TOROW", "_xlfn.UNICHAR", "_xlfn.UNICODE", "_xlfn.UNIQUE", "_xlfn.VAR.P", "_xlfn.VAR.S",
        "_xlfn.VSTACK", "_xlfn.WEBSERVICE", "_xlfn.WEIBULL.DIST", "_xlfn.WORKDAY.INTL", "_xlfn.WRAPCOLS",
        "_xlfn.WRAPROWS", "_xlfn.XLOOKUP", "_xlfn.XOR", "_xlfn.Z.TEST", "_xlfn._xlws.FILTER", "_xlfn._xlws.PY",
        "_xlfn._xlws.SORT"
    }.ToDictionary(name => name, name => name, StringComparer.OrdinalIgnoreCase);

    /// <summary>Returns the table spelling of a future-function name, or false when the name is not one.</summary>
    internal static bool TryGetName(string name, out string canonical) {
        if (Names.TryGetValue(name, out string? found)) {
            canonical = found;
            return true;
        }
        canonical = string.Empty;
        return false;
    }

    /// <summary>True for a future function whose plain call a native XLSB formula can encode.</summary>
    internal static bool IsEncodable(string canonical) => Names.ContainsKey(canonical) && !Excluded.Contains(canonical);

    /// <summary>True for a BrtName that is a future-function name with its fixed #NAME? formula.</summary>
    internal static bool IsFutureFunctionName(uint flags, byte[] formulaTokens) =>
        (flags & FutureFunctionFlag) != 0 && formulaTokens.Length == 2 && formulaTokens[0] == 0x1C && formulaTokens[1] == 0x1D;
}
