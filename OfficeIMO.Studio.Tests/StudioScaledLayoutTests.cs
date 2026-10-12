using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Tests;

/// <summary>
/// Common laptop screens at 125-200% display scaling leave Studio a small logical window. With the expanded pseudo-locale,
/// every visible command must stay inside the window (or its scroll area) and must not sit on top of another command.
/// </summary>
public sealed class StudioScaledLayoutTests {
    [Theory]
    [InlineData(1093, 614)] // 1366 x 768 at 125%
    [InlineData(853, 533)]  // 1280 x 800 at 150%
    [InlineData(960, 540)]  // 1920 x 1080 at 200%
    public async Task Commands_stay_reachable_and_separate_at_scaled_sizes_with_long_labels(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            IStudioLocalizer original = StudioLocalization.Current;
            StudioLocalization.Configure(new StudioLocalizer(CultureInfo.GetCultureInfo("en-XA")));
            var window = new MainWindow { Width = width, Height = height };
            var problems = new List<string>();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                MainWindowViewModel model = window.ViewModel;
                var screens = new (string Name, Action Show)[] {
                    ("View", () => model.ShowViewModeCommand.Execute(null)),
                    ("Annotate", () => model.ShowAnnotateModeCommand.Execute(null)),
                    ("Edit", () => model.ShowEditModeCommand.Execute(null)),
                    ("Pages", () => model.ShowPagesModeCommand.Execute(null)),
                    ("Forms", () => model.ShowFormsModeCommand.Execute(null)),
                    ("Protect", () => model.ShowProtectModeCommand.Execute(null)),
                    ("Home", () => model.ShowHomeCommand.Execute(null)),
                    ("Tools", () => model.ShowToolsCommand.Execute(null)),
                    ("Convert", () => model.ShowConversionWorkbenchCommand.Execute(null)),
                    ("OCR", () => model.ShowOcrCommand.Execute(null)),
                    ("Jobs", () => model.ShowJobsCommand.Execute(null)),
                    ("Settings", () => model.ShowSettingsCommand.Execute(null))
                };
                foreach ((string name, Action show) in screens) {
                    show();
                    Layout(window, width, height);
                    problems.AddRange(FindProblems(window).Select(problem => name + ": " + problem));
                    Capture(window, $"scaled-{width}x{height}-{name}");
                }
            } finally {
                window.Close();
                StudioLocalization.Configure(original);
            }
            Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems.Distinct().Take(40)));
            return true;
        }, CancellationToken.None);
    }

    private static IEnumerable<string> FindProblems(Window window) {
        Control[] commands = window.GetVisualDescendants().OfType<Control>()
            .Where(control => control is Button or ToggleButton or ComboBox or TextBox && control.IsEffectivelyVisible &&
                control.Bounds.Width > 0 && control.Bounds.Height > 0 && !IsInsidePopupOrTemplate(control))
            .ToArray();
        foreach (Control command in commands) {
            if (command.TranslatePoint(default, window) is not { } origin) continue;
            var bounds = new Rect(origin, command.Bounds.Size);
            ScrollViewer? scroller = command.FindAncestorOfType<ScrollViewer>();
            Rect area = scroller?.TranslatePoint(default, window) is { } scrollOrigin
                ? new Rect(scrollOrigin, scroller.Bounds.Size) : new Rect(window.Bounds.Size);
            bool scrollsHorizontally = scroller?.HorizontalScrollBarVisibility is ScrollBarVisibility.Auto or ScrollBarVisibility.Visible;
            if (!scrollsHorizontally && (bounds.Left < area.Left - 2 || bounds.Right > area.Right + 2))
                yield return $"{Describe(command)} extends past the {(scroller is null ? "window" : "scroll area")} horizontally ({bounds} in {area})";
            if (scroller is null && (bounds.Top < -2 || bounds.Bottom > window.Bounds.Height + 2))
                yield return $"{Describe(command)} is outside the window vertically ({bounds})";
        }
        foreach (IGrouping<Visual?, Control> siblings in commands.GroupBy(command => command.GetVisualParent())) {
            Control[] row = siblings.ToArray();
            for (int i = 0; i < row.Length; i++)
                for (int j = i + 1; j < row.Length; j++)
                    if (row[i].Bounds.Intersect(row[j].Bounds) is { Width: > 2, Height: > 2 })
                        yield return $"{Describe(row[i])} overlaps {Describe(row[j])}";
        }
    }

    private static bool IsInsidePopupOrTemplate(Control control) =>
        control.TemplatedParent is ComboBox or NumericUpDown or ScrollBar or Expander or Avalonia.Controls.Calendar or DatePicker or ButtonSpinner;

    private static string Describe(Control control) {
        string text = control switch {
            ContentControl { Content: string content } => content,
            TextBox { PlaceholderText: { } placeholder } => placeholder,
            _ => Avalonia.Automation.AutomationProperties.GetName(control) ?? control.Name ?? string.Empty
        };
        return control.GetType().Name + (text.Length > 0 ? " '" + (text.Length > 40 ? text[..40] + "…" : text) + "'" : string.Empty);
    }

    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        using var bitmap = Avalonia.Headless.HeadlessWindowExtensions.CaptureRenderedFrame(window);
        if (bitmap is null) return;
        Directory.CreateDirectory(output);
        bitmap.Save(Path.Combine(output, name + ".png"));
    }

    private static void Layout(Window window, int width, int height) {
        for (int pass = 0; pass < 3; pass++) {
            Dispatcher.UIThread.RunJobs();
            window.Measure(new Size(width, height));
            window.Arrange(new Rect(0, 0, width, height));
            window.UpdateLayout();
        }
        Dispatcher.UIThread.RunJobs();
    }
}
