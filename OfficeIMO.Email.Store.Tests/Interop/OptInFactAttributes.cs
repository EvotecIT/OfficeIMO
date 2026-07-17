using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EnvironmentFactAttribute : FactAttribute {
    public EnvironmentFactAttribute(string variableName, string? expectedValue = null,
        bool requireDirectory = false) {
        string? value = Environment.GetEnvironmentVariable(variableName);
        bool available = expectedValue == null
            ? !string.IsNullOrWhiteSpace(value)
            : string.Equals(value, expectedValue, StringComparison.Ordinal);
        if (available && requireDirectory) available = Directory.Exists(value);
        if (!available) {
            Skip = expectedValue == null
                ? string.Concat("Set ", variableName, " to run this opt-in test.")
                : string.Concat("Set ", variableName, "=", expectedValue,
                    " to run this opt-in test.");
        }
    }
}

public sealed class OutlookInteropFactAttribute : FactAttribute {
    public OutlookInteropFactAttribute() {
        if (!string.Equals(Environment.GetEnvironmentVariable(
            "OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP"), "1", StringComparison.Ordinal)) {
            Skip = "Set OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP=1 to run classic Outlook interoperability.";
        } else if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Skip = "Classic Outlook interoperability requires Windows.";
        } else {
#pragma warning disable CA1416
            if (Type.GetTypeFromProgID("Outlook.Application") == null) {
                Skip = "Classic Outlook is not registered on this machine.";
            }
#pragma warning restore CA1416
        }
    }
}
