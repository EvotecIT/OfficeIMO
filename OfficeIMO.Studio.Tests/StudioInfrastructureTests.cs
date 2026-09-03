using System.Globalization;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioInfrastructureTests {
    [Fact]
    public void PreferencesRoundTripVersionedCultureAndTheme() {
        using var folder = new TemporaryFolder("preferences");
        string path = Path.Combine(folder.Path, "preferences.json");
        var store = new JsonStudioPreferencesStore(path);

        store.Save(new StudioPreferences { UiCulture = "pl", Theme = StudioThemePreference.Dark });
        StudioPreferences loaded = store.Load();

        Assert.Equal(StudioPreferences.CurrentSchemaVersion, loaded.SchemaVersion);
        Assert.Equal("pl", loaded.UiCulture);
        Assert.Equal(StudioThemePreference.Dark, loaded.Theme);
        Assert.False(File.Exists(path + ".tmp"));
    }

    [Fact]
    public void PreferencesPreserveExplicitHighContrastTheme() {
        using var folder = new TemporaryFolder("preferences-high-contrast");
        var store = new JsonStudioPreferencesStore(Path.Combine(folder.Path, "preferences.json"));

        store.Save(new StudioPreferences { Theme = StudioThemePreference.HighContrast });

        Assert.Equal(StudioThemePreference.HighContrast, store.Load().Theme);
    }

    [Fact]
    public void PreferencesRecoverFromInvalidJsonAndCulture() {
        using var folder = new TemporaryFolder("preferences-invalid");
        string path = Path.Combine(folder.Path, "preferences.json");
        File.WriteAllText(path, "{broken");
        Assert.Equal(new StudioPreferences(), new JsonStudioPreferencesStore(path).Load());

        var normalized = new StudioPreferences { UiCulture = "not-a-culture" }.Normalize();
        Assert.Equal(StudioCultureCatalog.DefaultCulture, normalized.UiCulture);
    }

    [Fact]
    public void PseudolocaleExpandsTextWithoutChangingFormatItems() {
        var localizer = new StudioLocalizer(CultureInfo.GetCultureInfo(StudioCultureCatalog.PseudoCulture));

        string value = localizer.Get("Document.PagePosition");
        string formatted = localizer.Format("Document.PagePosition", 2, 8);

        Assert.StartsWith("⟦", value, StringComparison.Ordinal);
        Assert.Contains("{0}", value, StringComparison.Ordinal);
        Assert.Contains("{1}", value, StringComparison.Ordinal);
        Assert.Contains("2", formatted, StringComparison.Ordinal);
        Assert.Contains("8", formatted, StringComparison.Ordinal);
        Assert.EndsWith("···⟧", value, StringComparison.Ordinal);
    }

    [Fact]
    public void CultureCatalogKeepsUnreviewedPacksOutOfThePicker() {
        Assert.Contains(StudioCultureCatalog.Available, culture => culture.Name == "en");
        Assert.Contains(StudioCultureCatalog.Available, culture => culture.Name == StudioCultureCatalog.PseudoCulture);
        Assert.DoesNotContain(StudioCultureCatalog.Available, culture => culture.Name == "pl");
        Assert.True(StudioCultureCatalog.Planned.Count >= 15);
        Assert.Contains(StudioCultureCatalog.Planned, culture => culture.Name == "pl");
        Assert.Contains(StudioCultureCatalog.Planned, culture => culture.Name == "de");
        Assert.Contains(StudioCultureCatalog.Planned, culture => culture.Name == "fr");
        Assert.Contains(StudioCultureCatalog.Planned, culture => culture.Name == "it");
    }

    [Fact]
    public void DiagnosticsDoNotRecordExceptionMessagesOrSourcePaths() {
        using var folder = new TemporaryFolder("diagnostics");
        var diagnostics = new StudioDiagnostics(folder.Path);
        Exception exception;
        try {
            ThrowSensitiveFailure();
            throw new InvalidOperationException("unreachable");
        } catch (Exception caught) {
            exception = caught;
        }

        diagnostics.Write(StudioDiagnosticLevel.Error, "startup", "test-failure", exception);
        string log = File.ReadAllText(Path.Combine(folder.Path, "studio.log.jsonl"));

        Assert.DoesNotContain("confidential-document.pdf", log, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(nameof(StudioInfrastructureTests) + ".cs", log, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("test-failure", log, StringComparison.Ordinal);
        Assert.Contains("ExceptionType", log, StringComparison.Ordinal);
    }

    private static void ThrowSensitiveFailure() =>
        throw new InvalidOperationException("Failed to process C:\\Customers\\confidential-document.pdf");

    private sealed class TemporaryFolder : IDisposable {
        internal TemporaryFolder(string purpose) {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officeimo-studio-{purpose}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
