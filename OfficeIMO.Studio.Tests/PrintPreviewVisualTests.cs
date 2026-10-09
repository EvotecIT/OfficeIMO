using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class PrintPreviewVisualTests {
    [Theory]
    [InlineData(960, 640, 6, false)]
    [InlineData(1280, 850, 9, true)]
    [InlineData(720, 540, 6, false)]
    public async Task ExtendedPrintControlsRenderAndPrepareTheChosenLayout(int width, int height, int nUp, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "print-layout.pdf");
            PdfDocument.Create(document => {
                for (int number = 1; number <= 18; number++) {
                    int pageNumber = number;
                    document.Page(page => page.Size(200, 300).Margin(12).Background(PdfColor.LightGray)
                        .DefaultTextStyle(style => style.FontSize(36)).Content(content => content.Item(item =>
                        item.Paragraph(text => text.Text("Page " + pageNumber)))));
                }
            }).Save(source);
            byte[] original = File.ReadAllBytes(source);
            string? visualOutput = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
            if (!string.IsNullOrWhiteSpace(visualOutput)) {
                Directory.CreateDirectory(visualOutput);
                File.WriteAllBytes(Path.Combine(visualOutput, "print-layout-source.pdf"), original);
            }
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.ShowPrintPreviewCommand.Execute(null);
                var print = window.ViewModel.OutputWorkbench.PrintPreview;
                print.Pages = "18,17,16,15,14,13,12,11,10,9,8,7,6,5,4,3,2,1";
                window.UpdateLayout();
                ComboBox Choice(object choices) => window.GetVisualDescendants().OfType<ComboBox>()
                    .Single(control => ReferenceEquals(control.ItemsSource, choices));
                Choice(print.PageSubsetChoices).SelectedItem = print.PageSubsetChoices.Single(c => c.Value == PdfPrintPageSubset.Odd);
                Choice(print.PagesPerSheetChoices).SelectedItem = print.PagesPerSheetChoices.Single(c => c.Value == nUp);
                Choice(print.ScaleChoices).SelectedItem = print.ScaleChoices.Single(c => c.Value == PdfPrintScaleMode.Custom);
                Choice(print.AlignmentChoices).SelectedItem = print.AlignmentChoices.Single(c => c.Value == PdfPrintAlignment.BottomRight);
                Choice(print.ColorChoices).SelectedItem = print.ColorChoices.Single(c => c.Value == PdfPrintColorMode.Grayscale);
                var margins = window.GetVisualDescendants().OfType<Expander>().Single(control => Equals(control.Header, "Margins (points)"));
                margins.IsExpanded = true;
                window.UpdateLayout();
                NumericUpDown Number(string name) => window.GetVisualDescendants().OfType<NumericUpDown>()
                    .Single(control => AutomationProperties.GetName(control) == name);
                Number("Custom scale (%)").Value = 55;
                Number("Left").Value = 24; Number("Top").Value = 30;
                Number("Right").Value = 36; Number("Bottom").Value = 42;
                margins.BringIntoView(); window.UpdateLayout();
                Capture(window, "print-layout-controls-" + width);
                Assert.True(print.UsesCustomScale);
                Assert.Equal(55, print.CustomScalePercent);
                await print.BuildPreviewCommand.ExecuteAsync(null);
                Assert.True(print.HasPreview, print.Status);
                Assert.Equal(nUp == 6 ? 2 : 1, print.Sheets.Count);
                Assert.Equal(new[] { 17, 15, 13, 11, 9, 7, 5, 3, 1 }, print.Sheets.SelectMany(s => s.Placements.Select(p => p.PageNumber)));
                Assert.All(print.Sheets.SelectMany(s => s.Placements), p => {
                    Assert.Equal(0.55, p.Scale); Assert.Equal(110, p.Width, 5); Assert.Equal(165, p.Height, 5);
                    Assert.Equal(p.SlotX + p.SlotWidth - p.Width, p.X, 5);
                    Assert.Equal(p.SlotY + p.SlotHeight - p.Height, p.Y, 5);
                });
                TabControl? compactTabs = window.GetVisualDescendants().OfType<TabControl>().FirstOrDefault(control =>
                    control.Classes.Contains("pageTabs") && control.Items.OfType<TabItem>().Any(item => Equals(item.Header, "Details")));
                if (compactTabs is not null) {
                    compactTabs.SelectedIndex = 1;
                    window.UpdateLayout();
                }
                Capture(window, "print-layout-sheets-" + width);
                if (compactTabs is not null) { compactTabs.SelectedIndex = 0; window.UpdateLayout(); }
                Number("Left").Value = 25;
                Assert.False(print.HasPreview);
                Assert.Equal(original, File.ReadAllBytes(source));
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(960, 640, false)]
    [InlineData(1280, 850, true)]
    public async Task ReviewedSheetsUseUnsavedWorkspaceAndExposePrinterControls(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "print-current.pdf");
            PdfDocument.Create(document => {
                document.Page(page => page.Size(300, 400).Content(content => content.Item(item => item.Paragraph(text => text.Text("First reviewed page")))));
                document.Page(page => page.Size(300, 400).Content(content => content.Item(item => item.Paragraph(text => text.Text("Second reviewed page")))));
            }).Save(source);
            byte[] original = File.ReadAllBytes(source);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                var model = window.ViewModel;
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                Assert.Equal(3, model.Pages.Count);
                model.ShowPrintPreviewCommand.Execute(null);
                var print = model.OutputWorkbench.PrintPreview;
                print.SelectedPaper = print.PaperChoices.Single(choice => choice.Name == "A3");
                print.PrintDpi = 300;
                print.SelectedOrientation = print.OrientationChoices.Single(choice => choice.Value == PdfPrintOrientation.Landscape);
                print.SelectedPagesPerSheet = print.PagesPerSheetChoices.Single(choice => choice.Value == 2);
                await print.BuildPreviewCommand.ExecuteAsync(null);
                Assert.True(print.HasPreview, print.Status);
                Assert.Equal(2, print.Sheets.Count);
                Assert.Equal(3, print.Sheets.Sum(sheet => sheet.Placements.Count));
                string? nativePrinter = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_TEST_PRINT_QUEUE");
                if (string.IsNullOrWhiteSpace(nativePrinter)) nativePrinter = null;
                if (string.IsNullOrWhiteSpace(nativePrinter)) print.PrinterChoices = [new("Review file printer", true, true)];
                else await print.RefreshPrintersCommand.ExecuteAsync(null);
                print.SelectedPrinter = print.PrinterChoices.Single(printer => printer.Name == (nativePrinter ?? "Review file printer"));
                await print.PaperSourceDiscovery;
                Assert.False(print.CanPrint);
                print.PrintOutputPath = Path.Combine(services.Paths.Root, "reviewed-output.pdf");
                Assert.True(print.CanPrint);
                window.UpdateLayout();
                var sourceChoice = window.GetVisualDescendants().OfType<ComboBox>().Single(control => ReferenceEquals(control.ItemsSource, print.PaperSourceChoices));
                sourceChoice.BringIntoView();
                window.UpdateLayout();
                Assert.True(sourceChoice.IsEffectivelyEnabled);
                if (string.IsNullOrWhiteSpace(nativePrinter)) {
                    print.PaperSourceChoices = [new(null, "Printer default"), new("tray-2", "Lower tray")];
                }
                if (print.PaperSourceChoices.Count > 1) {
                    sourceChoice.SelectedItem = print.PaperSourceChoices[1];
                    Assert.NotNull(print.SelectedPaperSource?.Id);
                } else Assert.Null(print.SelectedPaperSource?.Id);
                Capture(window, "print-paper-source-" + width);
                Capture(window, "print-sheets-" + width);
                var printButton = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, print.PrintCommand));
                printButton.BringIntoView();
                window.UpdateLayout();
                Assert.True(printButton.IsEffectivelyEnabled);
                Point? location = printButton.TranslatePoint(new Point(), window);
                Assert.NotNull(location);
                Assert.InRange(location.Value.Y, 0, height - printButton.Bounds.Height);
                Capture(window, "print-delivery-controls-" + width);
                if (!string.IsNullOrWhiteSpace(nativePrinter)) {
                    await print.PrintCommand.ExecuteAsync(null);
                    Assert.StartsWith("Printer accepted job", print.Status);
                    await TextEditingReviewTests.WaitUntilAsync(() => File.Exists(print.PrintOutputPath) && new FileInfo(print.PrintOutputPath).Length > 0);
                    PdfDocument delivered = PdfDocument.Load(print.PrintOutputPath);
                    Assert.Equal(2, delivered.GetPageLayouts().Count);
                    Capture(window, "print-accepted-" + width);
                }
                Assert.Equal(original, File.ReadAllBytes(source));
                print.Status = "Printer accepted job queue-42: 2 sheet(s), 1 copy/copies. Check the printer for completion. " +
                    "Could not remove print staging at '/tmp/officeimo-print-private-staging': Access denied.";
                print.PrinterDiscoveryError = "Printer discovery failed: driver offline. The accepted job receipt remains available.";
                window.UpdateLayout();
                Capture(window, "print-accepted-cleanup-warning-" + width);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.RotateRightCommand.ExecuteAsync(null);
                Assert.False(print.HasPreview);
                Assert.False(print.CanPrint);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    private static void Capture(Window window, string name) {
        using var image = window.CaptureRenderedFrame();
        Assert.NotNull(image);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        image.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
