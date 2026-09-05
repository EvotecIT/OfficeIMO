using Avalonia;
using Avalonia.Headless;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

internal static class TestAppBuilder {
    private static readonly string TestRoot = Path.Combine(Path.GetTempPath(), "officeimo-studio-headless-" + Guid.NewGuid().ToString("N"));

    static TestAppBuilder() {
        AppDomain.CurrentDomain.ProcessExit += (_, _) => {
            try { if (Directory.Exists(TestRoot)) Directory.Delete(TestRoot, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        };
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder
            .Configure(() => new App(StudioApplicationServices.Create(
                new StudioDataPaths(Path.Combine(TestRoot, Guid.NewGuid().ToString("N"))))))
            .UseHeadless(new AvaloniaHeadlessPlatformOptions());

    internal static HeadlessUnitTestSession StartSession() =>
        HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder), AvaloniaTestIsolationLevel.PerTest);
}
