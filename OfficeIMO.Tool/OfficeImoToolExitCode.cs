namespace OfficeIMO.Tool;

internal enum OfficeImoToolExitCode {
    Success = 0,
    ValidationFailed = 1,
    Usage = 2,
    InputNotFound = 3,
    UnsupportedInput = 4,
    OperationFailed = 5,
    OutputFailed = 6,
    Cancelled = 130
}

internal static class OfficeImoToolExitCodes {
    internal static string Describe(OfficeImoToolExitCode code) => code switch {
        OfficeImoToolExitCode.Success => "The command completed.",
        OfficeImoToolExitCode.ValidationFailed => "The requested validation ran and found errors, for example invoice validate, markup validate or provenance check.",
        OfficeImoToolExitCode.Usage => "The command or an option is invalid, including a conversion with no matching route.",
        OfficeImoToolExitCode.InputNotFound => "An input file or folder was not found.",
        OfficeImoToolExitCode.UnsupportedInput => "The input is unsupported or could not be read.",
        OfficeImoToolExitCode.OperationFailed => "The document operation failed, including batches with failed items.",
        OfficeImoToolExitCode.OutputFailed => "The output could not be written, including a refused overwrite or an output over its size limit.",
        OfficeImoToolExitCode.Cancelled => "The command was cancelled.",
        _ => throw new ArgumentOutOfRangeException(nameof(code), code, "The exit code has no description.")
    };
}
