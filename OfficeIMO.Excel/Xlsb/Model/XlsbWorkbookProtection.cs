namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents the classic workbook protection verifier and flags.</summary>
    internal sealed class XlsbWorkbookProtection {
        internal XlsbWorkbookProtection(ushort workbookPassword, ushort revisionsPassword, ushort flags) {
            WorkbookPassword = workbookPassword;
            RevisionsPassword = revisionsPassword;
            Flags = flags;
        }

        internal ushort WorkbookPassword { get; }

        internal ushort RevisionsPassword { get; }

        internal ushort Flags { get; }

        internal bool LockStructure => (Flags & 0x0001) != 0;

        internal bool LockWindows => (Flags & 0x0002) != 0;

        internal bool LockRevision => (Flags & 0x0004) != 0;

        internal bool IsEmpty => WorkbookPassword == 0 && RevisionsPassword == 0 && Flags == 0;
    }
}
