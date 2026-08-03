using Xunit;

namespace OfficeIMO.Tests {
    internal sealed class VbaOfficeSipInteropFactAttribute : FactAttribute {
        public VbaOfficeSipInteropFactAttribute() {
            if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VBA_SIP_INTEROP"),
                    "1", StringComparison.Ordinal)) {
                Skip = "Set OFFICEIMO_RUN_VBA_SIP_INTEROP=1 to run Microsoft OfficeSips interoperability.";
            } else if (!System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(
                           System.Runtime.InteropServices.OSPlatform.Windows)) {
                Skip = "Microsoft OfficeSips interoperability requires Windows.";
            } else if (!File.Exists(Environment.GetEnvironmentVariable("OFFICEIMO_VBA_INTEROP_DOCUMENT_PATH"))) {
                Skip = "OFFICEIMO_VBA_INTEROP_DOCUMENT_PATH must identify a real macro-enabled Office document.";
            }
        }
    }
}
