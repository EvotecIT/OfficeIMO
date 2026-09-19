using System.Reflection;
using QuestPDF.Infrastructure;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class QuestPdfBenchmarkPolicy {
    internal const string PackageVersion = "2026.5.0";
    internal const string License = "MIT";
    internal const string LicenseUrl = "https://github.com/QuestPDF/QuestPDF/blob/2026.5.0/LICENSE.md";

    internal static string ConfiguredPackageVersion =>
        typeof(QuestPdfBenchmarkPolicy).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(static attribute => attribute.Key == "QuestPdfBenchmarkVersion")?.Value
        ?? PackageVersion;

    internal static bool InternalAuthorization =>
        bool.TryParse(
            typeof(QuestPdfBenchmarkPolicy).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(static attribute => attribute.Key == "QuestPdfInternalAuthorization")?.Value,
            out bool authorized) && authorized;

    internal static string LoadedAssemblyVersion =>
        typeof(IDocument).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(IDocument).Assembly.GetName().Version?.ToString()
        ?? "unknown";

    internal static string ConfiguredLicense =>
        string.Equals(ConfiguredPackageVersion, PackageVersion, StringComparison.Ordinal)
            ? License
            : "See selected QuestPDF package license";

    internal static string ConfiguredLicenseUrl =>
        $"https://github.com/QuestPDF/QuestPDF/blob/{ConfiguredPackageVersion}/LICENSE.md";

    internal static string ConfiguredLicenseType =>
        Environment.GetEnvironmentVariable("OFFICEIMO_QUESTPDF_LICENSE_TYPE") ?? nameof(LicenseType.Community);

    internal static void ConfigureLicense() {
        string configuredVersion = ConfiguredPackageVersion;
        if (!string.Equals(configuredVersion, PackageVersion, StringComparison.Ordinal) && !InternalAuthorization) {
            throw new InvalidOperationException(
                $"QuestPDF {configuredVersion} requires the guarded internal benchmark authorization path.");
        }
        if (!Version.TryParse(configuredVersion, out Version? expectedVersion)) {
            throw new InvalidOperationException($"QuestPDF benchmark package version '{configuredVersion}' is invalid.");
        }
        Version? version = typeof(IDocument).Assembly.GetName().Version;
        if (version is null ||
            version.Major != expectedVersion.Major ||
            version.Minor != expectedVersion.Minor ||
            version.Build != expectedVersion.Build) {
            throw new InvalidOperationException(
                $"QuestPDF benchmark policy requires package {configuredVersion}, but assembly {version?.ToString() ?? "unknown"} was loaded.");
        }

        string configuredLicenseType = ConfiguredLicenseType;
        if (!Enum.TryParse(configuredLicenseType, ignoreCase: true, out LicenseType licenseType)) {
            throw new InvalidOperationException($"QuestPDF license type '{configuredLicenseType}' is invalid.");
        }
        QuestPDF.Settings.License = licenseType;
    }
}
