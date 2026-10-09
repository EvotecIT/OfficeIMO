using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
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
            timeout.Token.ThrowIfCancellationRequested();
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
            await Task.Delay(10, timeout.Token);
        }
    }
}
