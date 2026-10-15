using Avalonia;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio;

internal static class Program {
    private const string NewWindowArgument = "--new-window";

    [STAThread]
    public static void Main(string[] args) {
        bool independent = args.Contains(NewWindowArgument, StringComparer.OrdinalIgnoreCase);
        string[] documents = args.Where(argument => !string.Equals(argument, NewWindowArgument, StringComparison.OrdinalIgnoreCase)).ToArray();
        using StudioSingleInstance? instance = independent ? null : AcquireInstance();
        if (instance is { IsPrimary: false } && instance.TryForward(FullPaths(documents), TimeSpan.FromSeconds(3))) return;
        BuildAvaloniaApp(instance is { IsPrimary: true } ? instance : null).StartWithClassicDesktopLifetime(documents);
    }

    internal static AppBuilder BuildAvaloniaApp() => BuildAvaloniaApp(null);

    private static AppBuilder BuildAvaloniaApp(StudioSingleInstance? instance) =>
        AppBuilder
            .Configure(() => new App { SingleInstance = instance })
            .UsePlatformDetect()
            .LogToTrace();

    // macOS delivers later documents to the running application through LaunchServices activation.
    private static StudioSingleInstance? AcquireInstance() =>
        OperatingSystem.IsWindows() || OperatingSystem.IsLinux() || OperatingSystem.IsFreeBSD()
            ? StudioSingleInstance.Acquire(StudioSingleInstance.CreateName(StudioDataPaths.CreateDefault().Root))
            : null;

    private static string[] FullPaths(IEnumerable<string> arguments) => arguments
        .Where(static argument => !string.IsNullOrWhiteSpace(argument))
        .Select(static argument => {
            try { return Path.GetFullPath(argument); }
            catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException) { return argument; }
        }).ToArray();
}
