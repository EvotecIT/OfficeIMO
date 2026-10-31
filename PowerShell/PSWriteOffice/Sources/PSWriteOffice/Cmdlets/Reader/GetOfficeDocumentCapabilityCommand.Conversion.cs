using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Reader;

public sealed partial class GetOfficeDocumentCapabilityCommand {
    private void WriteConversions() {
        OfficeWorkflowCommandUtilities.RegisterConversions();
        var executable = new HashSet<string>(OfficeConverter.Routes.Select(route => route.Id), StringComparer.OrdinalIgnoreCase);
        OfficeFormatSelection? from = OfficeFormatSelection.Parse(From);
        OfficeFormatSelection? to = OfficeFormatSelection.Parse(To);
        foreach (OfficeConversionCapability route in OfficeConversionCapabilityCatalog.All) {
            if (Id != null && !Id.Contains(route.Id, StringComparer.OrdinalIgnoreCase)) continue;
            if (from != null && !route.SourceExtensions.Any(from.Matches)) continue;
            if (to != null && !to.Matches(route.TargetExtension)) continue;
            bool supported = executable.Contains(route.Id);
            if (Executable && !supported) continue;
            var row = new PSObject(route);
            row.Properties.Add(new PSNoteProperty("ConversionSupported", supported));
            row.Properties.Add(new PSNoteProperty("ConversionCommand", supported ? "Convert-OfficeDocument" : null));
            WriteObject(row);
        }
    }
}
