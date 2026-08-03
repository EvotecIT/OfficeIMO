using Xunit;
using System.IO;

namespace OfficeIMO.Security.Tests;

internal sealed class VbaOfficeSipInteropFactAttribute : FactAttribute {
    public VbaOfficeSipInteropFactAttribute() {
        if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VBA_SIP_INTEROP"),
                "1", StringComparison.Ordinal)) {
            Skip = "Set OFFICEIMO_RUN_VBA_SIP_INTEROP=1 to run the Microsoft Office SIP differential corpus.";
        } else if (!OperatingSystem.IsWindows()) {
            Skip = "The optional Microsoft Office SIP differential requires Windows.";
        } else if (!Directory.Exists(Environment.GetEnvironmentVariable("OFFICEIMO_VBA_INTEROP_CORPUS"))) {
            Skip = "OFFICEIMO_VBA_INTEROP_CORPUS must identify the downloaded producer corpus.";
        }
    }
}
