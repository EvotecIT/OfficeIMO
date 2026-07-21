using Xunit;

namespace OfficeIMO.Tests;

internal sealed class LegacyPptComFactAttribute : FactAttribute {
    public LegacyPptComFactAttribute() {
        string? value = Environment.GetEnvironmentVariable("OFFICEIMO_RUN_LEGACY_PPT_COM_VALIDATION");
        if (!string.Equals(value, "1", StringComparison.Ordinal)
            && !string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) {
            Skip = "Set OFFICEIMO_RUN_LEGACY_PPT_COM_VALIDATION=1 to run desktop PowerPoint COM validation.";
        }
    }
}
