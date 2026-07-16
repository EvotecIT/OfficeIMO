namespace OfficeIMO.Reader.Tool;

internal enum ReaderToolExitCode {
    Success = 0,
    Usage = 2,
    InputNotFound = 3,
    UnsupportedInput = 4,
    ReadFailed = 5,
    OutputFailed = 6,
    Cancelled = 130
}