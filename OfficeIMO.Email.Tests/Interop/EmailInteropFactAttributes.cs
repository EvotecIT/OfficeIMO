using System.Runtime.InteropServices;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailArtifactOutlookInteropFactAttribute : FactAttribute {
    public EmailArtifactOutlookInteropFactAttribute() {
        if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_OUTLOOK_INTEROP"), "1",
                StringComparison.Ordinal)) {
            Skip = "Set OFFICEIMO_EMAIL_OUTLOOK_INTEROP=1 to run Outlook artifact interoperability.";
        } else if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Skip = "Outlook artifact interoperability requires Windows.";
        } else {
#pragma warning disable CA1416
            if (Type.GetTypeFromProgID("Outlook.Application") == null) {
                Skip = "Classic Outlook is not registered on this machine.";
            }
#pragma warning restore CA1416
        }
    }
}

public sealed class ExternalOutlookSmimeCorpusFactAttribute : FactAttribute {
    public ExternalOutlookSmimeCorpusFactAttribute() {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_SMIME_CORPUS");
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) {
            Skip = "Set OFFICEIMO_EMAIL_SMIME_CORPUS to the Outlook S/MIME corpus root to run this test.";
        }
    }
}
