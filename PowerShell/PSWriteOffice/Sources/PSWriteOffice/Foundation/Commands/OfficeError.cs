using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>Creates stable PowerShell errors whose target is the failing input.</summary>
public static class OfficeError {
    /// <summary>Creates an error for a diagnostic code. PowerShell appends the implementing cmdlet type itself.</summary>
    /// <remarks>Use the diagnostic code alone when writing through a cmdlet, so the runtime adds exactly one command suffix.</remarks>
    public static ErrorRecord Create(Exception exception, string diagnosticCode, object? input, ErrorCategory? category = null) {
        if (exception == null) throw new ArgumentNullException(nameof(exception));
        if (string.IsNullOrWhiteSpace(diagnosticCode)) throw new ArgumentException("A diagnostic code is required.", nameof(diagnosticCode));
        return new ErrorRecord(exception, diagnosticCode, category ?? Category(exception), input);
    }
    /// <summary>Maps a Core diagnostic to an error without losing its stable code.</summary>
    public static ErrorRecord Create(OfficeDiagnostic diagnostic, object? input) => Create(new InvalidOperationException(diagnostic.Message), diagnostic.Code, input);
    /// <summary>Classifies reachable argument, permission, I/O and cancellation failures.</summary>
    public static ErrorCategory Category(Exception exception) => exception switch {
        ArgumentException or FormatException => ErrorCategory.InvalidArgument,
        FileNotFoundException or DirectoryNotFoundException or ItemNotFoundException => ErrorCategory.ObjectNotFound,
        UnauthorizedAccessException or System.Security.SecurityException => ErrorCategory.PermissionDenied,
        OperationCanceledException or PipelineStoppedException => ErrorCategory.OperationStopped,
        IOException => ErrorCategory.WriteError,
        NotSupportedException => ErrorCategory.NotImplemented,
        _ => ErrorCategory.InvalidOperation
    };
}
