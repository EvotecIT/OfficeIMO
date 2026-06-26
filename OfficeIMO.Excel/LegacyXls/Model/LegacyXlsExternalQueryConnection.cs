namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only DBQueryExt metadata for an external query table or PivotCache connection.
    /// </summary>
    public sealed class LegacyXlsExternalQueryConnection {
        /// <summary>
        /// Creates decoded DBQueryExt metadata.
        /// </summary>
        public LegacyXlsExternalQueryConnection(
            int recordOffset,
            ushort recordType,
            string? sheetName,
            ushort futureRecordType,
            ushort dataSourceType,
            LegacyXlsExternalQueryConnectionSourceType sourceTypeKind,
            string sourceTypeName,
            ushort connectionFlags,
            ushort sourceSpecificFlags,
            ushort queryOptions,
            byte editVersion,
            byte refreshedVersion,
            byte refreshableMinimumVersion,
            ushort oleDbConnectionCount,
            ushort futureByteCount,
            ushort refreshIntervalMinutes,
            ushort htmlFormat,
            ushort parameterFlagCount,
            int parameterFlagByteCount,
            bool hasCompleteParameterFlags) {
            if (parameterFlagByteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(parameterFlagByteCount));
            }

            RecordOffset = recordOffset;
            RecordType = recordType;
            SheetName = sheetName;
            FutureRecordType = futureRecordType;
            DataSourceType = dataSourceType;
            SourceTypeKind = sourceTypeKind;
            SourceTypeName = sourceTypeName ?? throw new ArgumentNullException(nameof(sourceTypeName));
            ConnectionFlags = connectionFlags;
            SourceSpecificFlags = sourceSpecificFlags;
            QueryOptions = queryOptions;
            EditVersion = editVersion;
            RefreshedVersion = refreshedVersion;
            RefreshableMinimumVersion = refreshableMinimumVersion;
            OleDbConnectionCount = oleDbConnectionCount;
            FutureByteCount = futureByteCount;
            RefreshIntervalMinutes = refreshIntervalMinutes;
            HtmlFormat = htmlFormat;
            ParameterFlagCount = parameterFlagCount;
            ParameterFlagByteCount = parameterFlagByteCount;
            HasCompleteParameterFlags = hasCompleteParameterFlags;
        }

        /// <summary>Gets the byte offset of the DBQueryExt BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the worksheet or sheet entry name associated with the record, when known.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the FrtHeaderOld rt value embedded in the record payload.</summary>
        public ushort FutureRecordType { get; }

        /// <summary>Gets the raw DataSourceType value from the DBQueryExt record.</summary>
        public ushort DataSourceType { get; }

        /// <summary>Gets the decoded source type family when the value is known.</summary>
        public LegacyXlsExternalQueryConnectionSourceType SourceTypeKind { get; }

        /// <summary>Gets the decoded source type name, or a stable raw value for unknown source types.</summary>
        public string SourceTypeName { get; }

        /// <summary>Gets the raw DBQueryExt connection flag bits.</summary>
        public ushort ConnectionFlags { get; }

        /// <summary>Gets the raw ConnGrbitDbt source-specific flag bits.</summary>
        public ushort SourceSpecificFlags { get; }

        /// <summary>Gets the raw DBQueryExt query option flag bits.</summary>
        public ushort QueryOptions { get; }

        /// <summary>Gets whether the database connection remains open once established.</summary>
        public bool MaintainConnection => (ConnectionFlags & 0x0001) != 0;

        /// <summary>Gets whether the connection has not been refreshed.</summary>
        public bool NewQuery => (ConnectionFlags & 0x0002) != 0;

        /// <summary>Gets whether a Web query imports the underlying XML source instead of a page table.</summary>
        public bool ImportXmlSource => (ConnectionFlags & 0x0004) != 0;

        /// <summary>Gets whether the external connection uses the Web-based SharePoint list provider.</summary>
        public bool SharePointListSource => (ConnectionFlags & 0x0008) != 0;

        /// <summary>Gets whether SharePoint list data is reinitialized instead of refreshed.</summary>
        public bool SharePointListReinitializeCache => (ConnectionFlags & 0x0010) != 0;

        /// <summary>Gets whether the external connection source is XML.</summary>
        public bool SourceIsXml => (ConnectionFlags & 0x0080) != 0;

        /// <summary>Gets whether the DBQueryExt record is followed by a TxtQry record.</summary>
        public bool HasTextWizardQuery => (QueryOptions & 0x0001) != 0;

        /// <summary>Gets whether table names are stored in a following ExtString record.</summary>
        public bool HasTableNames => (QueryOptions & 0x0002) != 0;

        /// <summary>Gets the data functionality level that last edited the connection.</summary>
        public byte EditVersion { get; }

        /// <summary>Gets the data functionality level that last refreshed the connection.</summary>
        public byte RefreshedVersion { get; }

        /// <summary>Gets the minimum data functionality level required to refresh the connection.</summary>
        public byte RefreshableMinimumVersion { get; }

        /// <summary>Gets the number of OleDbConn records expected to follow this record.</summary>
        public ushort OleDbConnectionCount { get; }

        /// <summary>Gets the declared number of future-version bytes appended to the record.</summary>
        public ushort FutureByteCount { get; }

        /// <summary>Gets the automatic refresh interval in minutes. A value of zero disables timed refresh.</summary>
        public ushort RefreshIntervalMinutes { get; }

        /// <summary>Gets the Web query HTML formatting mode.</summary>
        public ushort HtmlFormat { get; }

        /// <summary>Gets the declared count of two-byte PBT parameter flag entries.</summary>
        public ushort ParameterFlagCount { get; }

        /// <summary>Gets the byte count occupied by PBT parameter flag entries in this record.</summary>
        public int ParameterFlagByteCount { get; }

        /// <summary>Gets whether the parameter flag byte count matches the declared PBT count.</summary>
        public bool HasCompleteParameterFlags { get; }

        /// <summary>Gets stable names for enabled DBQueryExt connection flags.</summary>
        public IReadOnlyList<string> ConnectionFlagNames {
            get {
                var names = new List<string>();
                if (MaintainConnection) names.Add("MaintainConnection");
                if (NewQuery) names.Add("NewQuery");
                if (ImportXmlSource) names.Add("ImportXmlSource");
                if (SharePointListSource) names.Add("SharePointListSource");
                if (SharePointListReinitializeCache) names.Add("SharePointListReinitializeCache");
                if (SourceIsXml) names.Add("SourceIsXml");
                return names;
            }
        }

        /// <summary>Gets stable names for enabled DBQueryExt query option flags.</summary>
        public IReadOnlyList<string> QueryOptionNames {
            get {
                var names = new List<string>();
                if (HasTextWizardQuery) names.Add("TextWizardQuery");
                if (HasTableNames) names.Add("TableNames");
                return names;
            }
        }
    }
}
