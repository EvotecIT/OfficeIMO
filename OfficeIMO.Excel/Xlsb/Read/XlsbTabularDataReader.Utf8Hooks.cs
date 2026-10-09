namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        partial void InitializeUtf8TextState();
        partial void ResetUtf8TextState();
        partial void TrackUtf8TextCell(int ordinal, int recordType);
        partial void TrackUtf8SharedStringCell(int ordinal, int sharedStringIndex);
        partial void ReleaseUtf8TextState();
    }
}
