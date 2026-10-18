using System.Management.Automation;

namespace PSWriteOffice.Cmdlets.Csv;

internal static class CsvCommandValidation
{
    public static void EnsureHeaderOptions(SwitchParameter noHeader, string[]? header)
    {
        if (noHeader.IsPresent && header is { Length: > 0 })
        {
            throw new PSArgumentException("Specify either -Header or -NoHeader, not both.");
        }
    }
}
