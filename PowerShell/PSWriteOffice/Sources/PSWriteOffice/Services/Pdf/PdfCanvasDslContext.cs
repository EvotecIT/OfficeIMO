using System;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Tracks the active page canvas while an existing-PDF canvas callback runs.</summary>
internal sealed class PdfCanvasDslContext : IDisposable
{
    private static readonly AsyncLocal<PdfCanvasDslContext?> CurrentScope = new();
    private readonly PdfCanvasDslContext? _previousScope;

    private PdfCanvasDslContext(PdfPageCanvas canvas, PdfStampPageContext page)
    {
        Canvas = canvas ?? throw new ArgumentNullException(nameof(canvas));
        Page = page ?? throw new ArgumentNullException(nameof(page));
        _previousScope = CurrentScope.Value;
        CurrentScope.Value = this;
    }

    /// <summary>Canvas receiving fixed-position content.</summary>
    public PdfPageCanvas Canvas { get; }

    /// <summary>Current target-page information.</summary>
    public PdfStampPageContext Page { get; }

    /// <summary>Enters a page-canvas DSL scope for the current callback.</summary>
    public static PdfCanvasDslContext Enter(PdfPageCanvas canvas, PdfStampPageContext page)
        => new(canvas, page);

    /// <summary>Returns the active canvas scope or throws a PowerShell-facing usage error.</summary>
    public static PdfCanvasDslContext Require(PSCmdlet caller)
        => CurrentScope.Value ?? throw new PSInvalidOperationException(
            $"No active PDF canvas context. Use {caller.MyInvocation.InvocationName} inside Add-OfficePdfCanvas -Content {{ ... }}.");

    /// <inheritdoc />
    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = _previousScope;
        }
    }
}
