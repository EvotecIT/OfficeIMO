using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using System.Globalization;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Tests;

/// <summary>A right-to-left interface language mirrors the workspace while documents keep their own direction.</summary>
public sealed class StudioRightToLeftLayoutTests {
    [Theory]
    [InlineData("en", 1280, 820)]
    [InlineData("ar", 1280, 820)]
    [InlineData("he", 960, 620)]
    public async Task TheWorkspaceMirrorsAndPagesKeepTheirDirection(string culture, int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            IStudioLocalizer original = StudioLocalization.Current;
            // English is the left-to-right baseline the mirrored layouts are compared with.
            CultureInfo rightToLeft = CultureInfo.GetCultureInfo(culture);
            App.ApplyFlowDirection(Application.Current!.Styles, rightToLeft);
            StudioLocalization.Configure(new StudioLocalizer(rightToLeft));
            var window = new MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                window.Width = width;
                window.Height = height;
                window.ApplyResponsiveLayout(width);
                window.Measure(new Size(width, height));
                window.Arrange(new Rect(0, 0, width, height));
                window.UpdateLayout();
                Dispatcher.UIThread.RunJobs();
                for (int wait = 0; wait < 200 && window.ViewModel.Pages.Any(page => page.IsRendering); wait++) {
                    await Task.Delay(25);
                    Dispatcher.UIThread.RunJobs();
                }
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();

                bool mirrored = rightToLeft.TextInfo.IsRightToLeft;
                Assert.Equal(mirrored ? FlowDirection.RightToLeft : FlowDirection.LeftToRight, window.FlowDirection);
                PdfPageCanvas[] pages = window.GetVisualDescendants().OfType<PdfPageCanvas>().ToArray();
                Assert.NotEmpty(pages);
                Assert.All(pages, page => Assert.Equal(FlowDirection.LeftToRight, page.FlowDirection));

                // The page list sits on the right, beside the pages, as a mirrored sidebar does.
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                Control pageList = window.ReaderPagesListControl;
                var navigation = workspace.FindControl<Grid>("NavigationPane")!;
                if (navigation.IsEffectivelyVisible) {
                    double navigationCenter = navigation.TranslatePoint(new Point(navigation.Bounds.Width / 2, 0), window)!.Value.X;
                    double pagesCenter = pageList.TranslatePoint(new Point(pageList.Bounds.Width / 2, 0), window)!.Value.X;
                    Assert.True(mirrored ? navigationCenter > pagesCenter : navigationCenter < pagesCenter, $"Navigation at {navigationCenter}, pages at {pagesCenter}");
                }

                // Toolbars wrap inside the window rather than running off the leading edge.
                foreach (WrapPanel panel in workspace.GetVisualDescendants().OfType<WrapPanel>().Where(panel => panel.IsEffectivelyVisible)) {
                    foreach (Control child in panel.Children.Where(child => child.IsVisible)) {
                        Point position = child.TranslatePoint(default, window)!.Value;
                        Assert.True(position.X >= -1 && position.X + child.Bounds.Width <= width + 1,
                            $"{child.GetType().Name} at {position.X} with width {child.Bounds.Width} leaves the {width} wide window.");
                    }
                }

                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"rtl-{culture}-{width}x{height}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally {
                window.Close();
                StudioLocalization.Configure(original);
            }
            return true;
        }, CancellationToken.None);
    }
}
