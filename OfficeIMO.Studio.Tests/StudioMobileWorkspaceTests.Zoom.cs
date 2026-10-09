using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileWorkspaceTests {
    [Theory]
    [InlineData(390, 844)]
    [InlineData(1024, 768)]
    public async Task PinchKeepsThePagePointUnderTheFingers(int width, int height) {
        await WithTouchReaderAsync(width, height, (window, view, controller) => {
            controller.Document.SetTouchZoom(1);
            Layout(window, width, height);
            var scroll = view.FindControl<ScrollViewer>("PageScroll")!;
            var canvas = scroll.GetVisualDescendants().OfType<PdfPageCanvas>().Single();
            scroll.Offset = new Vector(90, 160);
            Layout(window, width, height);
            Point origin = PagePoint(canvas, scroll);
            Assert.True(new Rect(scroll.Bounds.Size).Contains(origin));
            Capture(window, $"pinch-before-{width}");
            // Headless touch input pumps rendering between each contact update.
            Pinch(window, scroll, origin, 1.5);
            Layout(window, width, height);
            Layout(window, width, height);
            AssertPointNear(origin, PagePoint(canvas, scroll));
            Assert.Equal(1.5, controller.Document.Zoom);
            Capture(window, $"pinch-after-{width}");

            origin = PagePoint(canvas, scroll);
            Pinch(window, scroll, origin, 1 / 1.5);
            Layout(window, width, height);
            Layout(window, width, height);
            Assert.Equal(1, controller.Document.Zoom);
            AssertPointNear(origin, PagePoint(canvas, scroll));
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task PageChangeCancelsTheOldPinchAndItsQueuedScrollAdjustment() {
        await WithTouchReaderAsync(390, 844, (window, view, controller) => {
            controller.Document.SetTouchZoom(1);
            Layout(window, 390, 844);
            var scroll = view.FindControl<ScrollViewer>("PageScroll")!;
            scroll.RaiseEvent(new PinchEventArgs(1.25, new Point(150, 250)));
            controller.Document.SelectedPage = controller.Document.Pages[1];
            scroll.Offset = default;
            scroll.RaiseEvent(new PinchEventArgs(2, new Point(150, 250)));
            Layout(window, 390, 844);
            Layout(window, 390, 844);
            Assert.Equal(1.25, controller.Document.Zoom);
            Assert.Equal(default, scroll.Offset);
            scroll.RaiseEvent(new PinchEndedEventArgs());
            scroll.RaiseEvent(new PinchEventArgs(1.2, new Point(150, 250)));
            scroll.RaiseEvent(new PinchEndedEventArgs());
            Layout(window, 390, 844);
            Assert.Equal(1.5, controller.Document.Zoom);
            return Task.CompletedTask;
        });
    }

    [Theory]
    [InlineData("page")]
    [InlineData("document")]
    [InlineData("viewport")]
    [InlineData("detach")]
    [InlineData("focus")]
    public async Task ReaderChangeBeforeTheFirstPinchMoveRequiresFreshContacts(string change) {
        await WithTouchReaderAsync(390, 844, async (window, view, controller) => {
            controller.Document.SetTouchZoom(1);
            Layout(window, 390, 844);
            var scroll = view.FindControl<ScrollViewer>("PageScroll")!;
            Point center = scroll.TranslatePoint(new Point(160, 250), window)!.Value;
            using var first = window.TouchBegin(new Point(center.X - 20, center.Y), RawInputModifiers.None);
            using var second = window.TouchBegin(new Point(center.X + 20, center.Y), RawInputModifiers.None);
            switch (change) {
                case "page": controller.Document.SelectedPage = controller.Document.Pages[1]; break;
                case "document": await controller.OpenSampleAsync(); break;
                case "viewport": Layout(window, 844, 390); break;
                case "detach": window.Content = null; window.Content = view; break;
                case "focus": controller.Document.IsFocusReading = true; controller.Document.IsFocusReading = false; break;
            }
            int width = change == "viewport" ? 844 : 390;
            int height = change == "viewport" ? 390 : 844;
            Layout(window, width, height);
            double zoom = controller.Document.Zoom;
            window.TouchMove(first, new Point(center.X - 40, center.Y), RawInputModifiers.None);
            window.TouchMove(second, new Point(center.X + 40, center.Y), RawInputModifiers.None);
            Layout(window, width, height);
            Assert.Equal(zoom, controller.Document.Zoom);

            // Replacing one finger must not revive a cancelled gesture while the other is still down.
            window.TouchEnd(first, new Point(center.X - 40, center.Y), RawInputModifiers.None);
            using var replacement = window.TouchBegin(new Point(center.X - 20, center.Y), RawInputModifiers.None);
            window.TouchMove(replacement, new Point(center.X - 30, center.Y), RawInputModifiers.None);
            Assert.Equal(zoom, controller.Document.Zoom);
            window.TouchEnd(replacement, new Point(center.X - 30, center.Y), RawInputModifiers.None);
            window.TouchEnd(second, new Point(center.X + 40, center.Y), RawInputModifiers.None);
            Pinch(window, scroll, new Point(scroll.Bounds.Width / 2, scroll.Bounds.Height / 2), 1.2);
            Layout(window, width, height);
            Assert.Equal(Math.Round(zoom * 1.2, 2), controller.Document.Zoom);
        });
    }

    [Fact]
    public async Task PinchRespectsZoomLimitsAndRecentersASmallerPage() {
        await WithTouchReaderAsync(390, 844, (window, view, controller) => {
            var scroll = view.FindControl<ScrollViewer>("PageScroll")!;
            var origin = new Point(scroll.Bounds.Width / 2, scroll.Bounds.Height / 2);
            scroll.RaiseEvent(new PinchEventArgs(100, origin));
            scroll.RaiseEvent(new PinchEndedEventArgs());
            Layout(window, 390, 844);
            Layout(window, 390, 844);
            Assert.Equal(3, controller.Document.Zoom);
            scroll.RaiseEvent(new PinchEventArgs(0.001, origin));
            scroll.RaiseEvent(new PinchEndedEventArgs());
            Layout(window, 390, 844);
            Layout(window, 390, 844);
            Assert.Equal(0.25, controller.Document.Zoom);
            Assert.Equal(default, scroll.Offset);
            var canvas = scroll.GetVisualDescendants().OfType<PdfPageCanvas>().Single();
            Assert.True(new Rect(scroll.Bounds.Size).Contains(canvas.TranslatePoint(default, scroll)!.Value));
            Assert.True(new Rect(scroll.Bounds.Size).Contains(canvas.TranslatePoint(new Point(canvas.Bounds.Width, canvas.Bounds.Height), scroll)!.Value));
            return Task.CompletedTask;
        });
    }

    private static Point PagePoint(PdfPageCanvas canvas, ScrollViewer scroll) =>
        canvas.TranslatePoint(new Point(canvas.Bounds.Width * 0.45, canvas.Bounds.Height * 0.4), scroll)!.Value;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PinchDoesNotOpenAPointerContextMenuButIndependentRequestsStillWork(bool move) {
        await WithTouchReaderAsync(390, 844, (window, view, _) => {
            var scroll = view.FindControl<ScrollViewer>("PageScroll")!;
            PointerPressedEventArgs? press = null;
            window.AddHandler(InputElement.PointerPressedEvent, (_, args) => press = args, RoutingStrategies.Bubble, handledEventsToo: true);
            Point point = scroll.TranslatePoint(new Point(150, 250), window)!.Value;
            using var firstTouch = window.TouchBegin(point, RawInputModifiers.None);
            Point secondPoint = new(point.X + 40, point.Y);
            using var secondTouch = window.TouchBegin(secondPoint, RawInputModifiers.None);
            if (move) window.TouchMove(firstTouch, new Point(point.X - 10, point.Y), RawInputModifiers.None);
            window.TouchEnd(firstTouch, point, RawInputModifiers.None);
            window.TouchEnd(secondTouch, secondPoint, RawInputModifiers.None);
            var pointerRequest = new ContextRequestedEventArgs(Assert.IsType<PointerPressedEventArgs>(press));
            scroll.RaiseEvent(pointerRequest);
            Assert.True(pointerRequest.Handled);

            var keyboardRequest = new ContextRequestedEventArgs();
            scroll.RaiseEvent(keyboardRequest);
            Assert.False(keyboardRequest.Handled);
            using var nextTouch = window.TouchBegin(point, RawInputModifiers.None);
            window.TouchEnd(nextTouch, point, RawInputModifiers.None);
            var nextRequest = new ContextRequestedEventArgs(Assert.IsType<PointerPressedEventArgs>(press));
            scroll.RaiseEvent(nextRequest);
            Assert.False(nextRequest.Handled);
            return Task.CompletedTask;
        });
    }

    private static void AssertPointNear(Point expected, Point actual) =>
        Assert.InRange(new Vector(actual.X - expected.X, actual.Y - expected.Y).Length, 0, 2);

    private static void Pinch(Window window, ScrollViewer scroll, Point origin, double scale) {
        Point center = scroll.TranslatePoint(origin, window)!.Value;
        using var first = window.TouchBegin(new Point(center.X - 20, center.Y), RawInputModifiers.None);
        using var second = window.TouchBegin(new Point(center.X + 20, center.Y), RawInputModifiers.None);
        var firstEnd = new Point(center.X - 20 * scale, center.Y);
        var secondEnd = new Point(center.X + 20 * scale, center.Y);
        window.TouchMove(first, firstEnd, RawInputModifiers.None);
        window.TouchMove(second, secondEnd, RawInputModifiers.None);
        window.TouchEnd(first, firstEnd, RawInputModifiers.None);
        window.TouchEnd(second, secondEnd, RawInputModifiers.None);
    }

    private static async Task WithTouchReaderAsync(int width, int height,
        Func<Window, MobileWorkspaceView, MobileDocumentController, Task> action) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-touch-zoom-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask);
                var view = new MobileWorkspaceView();
                view.Connect(controller);
                var window = new Window { Content = view };
                try {
                    window.Show();
                    await controller.OpenSampleAsync();
                    Layout(window, width, height);
                    await controller.Document.SelectedPage!.EnsureRenderedAsync();
                    Layout(window, width, height);
                    var navigation = view.FindControl<SplitView>("ApplicationNavigation")!;
                    var pane = navigation.GetVisualDescendants().OfType<Control>().Single(control => control.Name == "PART_PaneRoot");
                    var scroll = view.FindControl<ScrollViewer>("PageScroll")!;
                    await StudioHeadlessInput.WaitForTargetAsync(window, scroll, () => Layout(window, width, height),
                        () => navigation.DisplayMode != SplitViewDisplayMode.Inline || !navigation.IsPaneOpen ||
                            pane.Bounds.Width == navigation.OpenPaneLength);
                    await action(window, view, controller);
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }
}
