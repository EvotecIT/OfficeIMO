using Avalonia;
using Avalonia.Headless;

namespace OfficeIMO.Studio.Tests;

internal static class TestAppBuilder {
    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder
            .Configure<App>()
            .UseHeadless(new AvaloniaHeadlessPlatformOptions());

    internal static HeadlessUnitTestSession StartSession() =>
        HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder), AvaloniaTestIsolationLevel.PerTest);
}
