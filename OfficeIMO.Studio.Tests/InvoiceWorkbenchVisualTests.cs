using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class InvoiceWorkbenchVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task EditingAndRenderingRemainAccessibleInScrollableCompactAndWideViews(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string input = Path.Combine(services.Paths.Root, "Invoice with a long descriptive source filename.xml");
            File.WriteAllBytes(input, InvoiceSample.Create());
            using var model = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(input), _ => Task.FromResult<string?>(null)) { InputPath = input };
            model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.EditSource);
            model.EditNumber = "EDITED-42"; model.EditIssueDate = "2026-09-30";
            var view = new InvoiceWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                var scroll = (ScrollViewer)view.Content!;
                Assert.Equal(0, scroll.Offset.X); Assert.True(scroll.Extent.Height > scroll.Viewport.Height);
                Capture(window, $"invoice-{width}-edit.png");
                await model.RunCommand.ExecuteAsync(null);
                Assert.True(model.HasOutput, model.Status);
                scroll.Offset = new Vector(0, scroll.Extent.Height);
                window.UpdateLayout();
                Capture(window, $"invoice-{width}-report.png");
                model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.RenderPresentationPdf);
                var presentation = view.GetVisualDescendants().OfType<Expander>().First(e => e.IsEffectivelyVisible);
                presentation.IsExpanded = true;
                scroll.Offset = new Vector(0, 0); window.UpdateLayout();
                Assert.All(model.AvailableTargets, t => Assert.Equal(OfficeIMO.Invoicing.InvoiceSyntax.Cii, t.Options.Syntax));
                Capture(window, $"invoice-{width}-presentation.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
