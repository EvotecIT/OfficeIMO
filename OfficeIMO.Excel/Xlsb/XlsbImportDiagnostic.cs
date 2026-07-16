namespace OfficeIMO.Excel.Xlsb {
    /// <summary>Identifies the severity of an XLSB import finding.</summary>
    public enum XlsbImportDiagnosticSeverity {
        /// <summary>Informational import or preservation finding.</summary>
        Information,
        /// <summary>Content was preserved but was not fully projected.</summary>
        Warning,
        /// <summary>The workbook could not be imported safely.</summary>
        Error
    }

    /// <summary>Describes one structured XLSB import finding.</summary>
    public sealed class XlsbImportDiagnostic {
        internal XlsbImportDiagnostic(
            string code,
            XlsbImportDiagnosticSeverity severity,
            string message,
            string? partName = null,
            int? recordType = null,
            long? recordOffset = null) {
            Code = code;
            Severity = severity;
            Message = message;
            PartName = partName;
            RecordType = recordType;
            RecordOffset = recordOffset;
        }

        /// <summary>Gets the stable machine-readable finding code.</summary>
        public string Code { get; }

        /// <summary>Gets the finding severity.</summary>
        public XlsbImportDiagnosticSeverity Severity { get; }

        /// <summary>Gets the human-readable finding description.</summary>
        public string Message { get; }

        /// <summary>Gets the related package part, when known.</summary>
        public string? PartName { get; }

        /// <summary>Gets the related BIFF12 record type, when known.</summary>
        public int? RecordType { get; }

        /// <summary>Gets the related record offset within its binary part, when known.</summary>
        public long? RecordOffset { get; }

        /// <inheritdoc />
        public override string ToString() => $"{Severity} {Code}: {Message}";
    }

    /// <summary>Describes one BIFF12 record retained for fidelity but not projected into the normal workbook model.</summary>
    public sealed class XlsbPreservedRecordInfo {
        internal XlsbPreservedRecordInfo(string partName, int recordType, long offset, int payloadLength) {
            PartName = partName;
            RecordType = recordType;
            Offset = offset;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the package part containing the record.</summary>
        public string PartName { get; }

        /// <summary>Gets the BIFF12 record type.</summary>
        public int RecordType { get; }

        /// <summary>Gets the record offset within the binary part.</summary>
        public long Offset { get; }

        /// <summary>Gets the record payload length.</summary>
        public int PayloadLength { get; }
    }
}
