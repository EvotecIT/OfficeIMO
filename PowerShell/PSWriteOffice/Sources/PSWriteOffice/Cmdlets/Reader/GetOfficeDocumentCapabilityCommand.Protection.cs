using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Security;

namespace PSWriteOffice.Cmdlets.Reader;

public sealed partial class GetOfficeDocumentCapabilityCommand {
    private void WriteProtection() {
        var catalog = OfficeProtectionCapabilityCatalog.Current;
        if (AsJson || AsMarkdown) {
            if (AsJson && AsMarkdown) throw new PSArgumentException("Specify only one of -AsJson or -AsMarkdown.");
            if (Id != null || !string.IsNullOrWhiteSpace(Format) || Kind.HasValue || IncompleteOnly)
                throw new PSArgumentException("Catalog text output cannot be combined with row filters.");
            WriteObject(AsJson ? catalog.ToJson() : catalog.ToMarkdown());
            return;
        }
        var capabilities = catalog.Capabilities.AsEnumerable();
        if (Id != null) capabilities = capabilities.Where(item => Id.Contains(item.Id, StringComparer.Ordinal));
        if (!string.IsNullOrWhiteSpace(Format)) capabilities = capabilities.Where(item => item.FormatId.IndexOf(Format!.Trim(), StringComparison.OrdinalIgnoreCase) >= 0);
        if (Kind.HasValue) capabilities = capabilities.Where(item => item.Kind == Kind.Value);
        if (IncompleteOnly) {
            var incomplete = catalog.IncompleteCapabilities.Select(item => item.Id).ToArray();
            capabilities = capabilities.Where(item => incomplete.Contains(item.Id, StringComparer.Ordinal));
        }
        WriteObject(capabilities.ToArray(), enumerateCollection: true);
    }
}
