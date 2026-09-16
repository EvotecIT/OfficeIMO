using System.Runtime.InteropServices;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailArtifactOutlookInteropFactAttribute : FactAttribute {
    public EmailArtifactOutlookInteropFactAttribute() {
        if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_OUTLOOK_INTEROP"), "1",
                StringComparison.Ordinal)) {
            Skip = "Set OFFICEIMO_EMAIL_OUTLOOK_INTEROP=1 to run Outlook artifact interoperability.";
        } else if (!IsWindowsPlatform()) {
            Skip = "Outlook artifact interoperability requires Windows.";
        } else if (GetOutlookApplicationTypeOnWindows() == null) {
            Skip = "Classic Outlook is not registered on this machine.";
        }
    }

#if NET5_0_OR_GREATER
    [SupportedOSPlatformGuard("windows")]
#endif
    private static bool IsWindowsPlatform() =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    private static Type? GetOutlookApplicationTypeOnWindows() =>
        Type.GetTypeFromProgID("Outlook.Application");
}

public sealed class ExternalOutlookSmimeCorpusFactAttribute : FactAttribute {
    public ExternalOutlookSmimeCorpusFactAttribute() {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_SMIME_CORPUS");
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) {
            Skip = "Set OFFICEIMO_EMAIL_SMIME_CORPUS to the Outlook S/MIME corpus root to run this test.";
        }
    }
}
