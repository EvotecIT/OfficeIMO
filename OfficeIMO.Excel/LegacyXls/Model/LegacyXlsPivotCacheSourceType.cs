namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the source-data kind declared for a legacy XLS PivotCache.
    /// </summary>
    public enum LegacyXlsPivotCacheSourceType {
        /// <summary>The PivotCache source is a worksheet range or defined range.</summary>
        Sheet = 0x0001,

        /// <summary>The PivotCache source is external data.</summary>
        External = 0x0002,

        /// <summary>The PivotCache source is one or more consolidation ranges.</summary>
        Consolidation = 0x0004,

        /// <summary>The PivotCache source is an application-managed temporary scenario structure.</summary>
        Scenario = 0x0010
    }
}
