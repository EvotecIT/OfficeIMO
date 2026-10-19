using System.Management.Automation;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
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

    /// <summary>Resolves a reader whose format settings were captured when its handlers were registered.</summary>
    protected OfficeDocumentReader ResolveReader(ReaderAllOptions? handlerOptions)
    {
        if (Reader != null)
        {
            if (handlerOptions != null)
            {
                throw new PSArgumentException(
                    "Format-specific switches cannot alter a caller-provided immutable OfficeIMO 3 reader. " +
                    "Configure the reader with New-OfficeDocumentReader -ReaderAllOptions, or omit -Reader.");
            }

            return Reader;
        }

        return handlerOptions == null
            ? ReaderCommandUtilities.Reader
            : ReaderCommandUtilities.CreateReader(handlerOptions);
    }

    /// <summary>Resolves mutually exclusive bounded and unlimited email-store settings.</summary>
    protected static int? ResolveStoreItemLimit(int? maxStoreItems, SwitchParameter allStoreItems)
    {
        if (allStoreItems.IsPresent && maxStoreItems.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxStoreItems or -AllStoreItems, not both.");
        }

        return allStoreItems.IsPresent ? int.MaxValue : maxStoreItems;
    }
}
