namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the data source family declared by a legacy XLS DBQueryExt record.
    /// </summary>
    public enum LegacyXlsExternalQueryConnectionSourceType {
        /// <summary>The DBQueryExt source type is not one of the currently decoded values.</summary>
        Unknown = 0,

        /// <summary>ODBC-based source.</summary>
        Odbc = 1,

        /// <summary>DAO-based source.</summary>
        Dao = 2,

        /// <summary>Web query source.</summary>
        Web = 4,

        /// <summary>OLE DB-based source.</summary>
        OleDb = 5,

        /// <summary>Text query source.</summary>
        Text = 6,

        /// <summary>ADO recordset source.</summary>
        Ado = 7
    }
}
