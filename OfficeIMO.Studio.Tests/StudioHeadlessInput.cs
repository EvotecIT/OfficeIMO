using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace OfficeIMO.Studio.Tests;

internal static class StudioHeadlessInput {
    public static async Task<Point> WaitForTargetAsync(Window window, Control target, Action layout, Func<bool>? ready = null) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        Rect? previousBounds = null;
        Point? previousPoint = null;
        int stableFrames = 0;
        // Headless input pumps rendering itself. Commit the hit-test snapshot and let
        // moving geometry settle before callers capture pointer or touch coordinates.
        while (true) {
            if (timeout.IsCancellationRequested) throw TargetTimeout(window, target);
            layout();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            Dispatcher.UIThread.RunJobs();
            Point? point = target.TranslatePoint(new Point(target.Bounds.Width / 2, target.Bounds.Height / 2), window);
            bool hitsTarget = point is { } center && window.InputHitTest(center) is Visual hit &&
                hit.GetSelfAndVisualAncestors().Contains(target);
            stableFrames = hitsTarget && (ready?.Invoke() ?? true) &&
                previousBounds == target.Bounds && previousPoint == point ? stableFrames + 1 : 0;
            if (stableFrames >= 3) return point!.Value;
            previousBounds = target.Bounds;
            previousPoint = point;
            try { await Task.Delay(10, timeout.Token); }
            catch (OperationCanceledException) when (timeout.IsCancellationRequested) {
                throw TargetTimeout(window, target);
            }
        }
    }

    private static TimeoutException TargetTimeout(Window window, Control target) {
        Point? point = target.TranslatePoint(new Point(target.Bounds.Width / 2, target.Bounds.Height / 2), window);
        var hit = point is { } center ? window.InputHitTest(center) as Control : null;
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            using var frame = window.CaptureRenderedFrame();
            frame?.Save(Path.Combine(output, $"input-timeout-{target.Name ?? target.GetType().Name}-{window.Bounds.Width}.png"),
                Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        return new TimeoutException($"Input target {target.Name ?? target.GetType().Name} did not become stable and reachable. "
            + $"Visible={target.IsEffectivelyVisible}; enabled={target.IsEffectivelyEnabled}; bounds={target.Bounds}; "
            + $"point={point}; hit={hit?.Name ?? hit?.GetType().Name ?? "none"}; window={window.Bounds}.");
    }
}
