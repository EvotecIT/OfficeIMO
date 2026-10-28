using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Base class for commands that can use the module reader or a caller-configured reader.</summary>
public abstract class OfficeDocumentReaderCommandBase : PSCmdlet
{
    /// <summary>Optional immutable OfficeIMO reader with caller-configured handlers and processors.</summary>
    [Parameter]
    public OfficeDocumentReader? Reader { get; set; }

    /// <summary>Gets the caller-provided reader or the module's fully configured reader.</summary>
    protected OfficeDocumentReader EffectiveReader => Reader ?? ReaderCommandUtilities.Reader;

}
