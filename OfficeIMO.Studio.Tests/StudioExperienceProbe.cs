using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

/// <summary>Repeatable native UI acceptance against synthetic documents and an isolated profile.</summary>
internal static class StudioExperienceProbe {
    internal static int Run(string root, string scenario, int width, int height, string culture, string theme) {
        if (scenario is not ("provenance" or "home" or "document" or "tabs" or "assistant" or "connections" or "ocr" or "convert" or "invoices" or "invoices-standards" or "watermark" or "watermark-image")
            || width < 960 || height < 620 || !Enum.TryParse(theme, true, out StudioThemePreference appearance)) return 2;
        root = Path.GetFullPath(root);
        Directory.CreateDirectory(root);
        var paths = new StudioDataPaths(Path.Combine(root, "profile"));
        new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences { UiCulture = culture, Theme = appearance, RememberSession = false });
        var services = StudioApplicationServices.Create(paths);
        string source = Path.Combine(root, "quarterly-review.pdf");
        var sample = PdfDocument.Create(document => document.Page(page => page.Content(content => {
            content.Text("Quarterly review");
            content.Text("The approved budget is 42,000 EUR. The delivery deadline is 30 September.");
            content.Text("Action: prepare the project summary. Owner: the delivery team.");
        })));
        if (scenario is "watermark" or "watermark-image") {
            sample = sample.Stamp.Watermark(new PdfWatermarkOptions {
                Text = "REVIEW COPY", X = 100, Y = 200, Width = 220, Height = 80, FontSize = 28, RotationDegrees = 0,
                ImageBytes = scenario == "watermark-image"
                    ? Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=") : null
            });
        }
        sample.Save(source);
        return AppBuilder.Configure(() => new App(services)).UsePlatformDetect().LogToTrace()
            .AfterSetup(builder => Dispatcher.UIThread.Post(async () => {
                try {
                    var lifetime = (IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!;
                    var window = (MainWindow)lifetime.MainWindow!;
                    window.Width = width; window.Height = height;
                    if (scenario is not ("home" or "invoices" or "invoices-standards")) await window.TabHost.OpenDocumentAsync(source);
                    if (scenario == "tabs") {
                        string second = Path.Combine(root, "delivery-notes.pdf");
                        PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Delivery notes: the review is complete.")))).Save(second);
                        await window.TabHost.OpenDocumentAsync(second);
                    }
                    if (scenario is "assistant" or "connections") window.ViewModel.ToggleAssistantCommand.Execute(null);
                    if (scenario == "ocr") window.ViewModel.ShowOcrCommand.Execute(null);
                    if (scenario == "provenance") {
                        string textSource = Path.Combine(root, "review.html");
                        await File.WriteAllTextAsync(textSource, "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>review\u202Ethis</body></html>");
                        var workspace = window.ViewModel.ProvenanceWorkbench;
                        workspace.InputPath = textSource; workspace.OutputFolder = root;
                        await workspace.AssessCommand.ExecuteAsync(null); workspace.RemoveReferences = true;
                        window.ViewModel.ShowProvenanceCommand.Execute(null);
                    }
                    if (scenario == "convert") window.ViewModel.Commands["Convert"].Execute(null);
                    if (scenario is "invoices" or "invoices-standards") {
                        string invoice = Path.Combine(root, "sample-invoice.xml");
                        File.WriteAllBytes(invoice, InvoiceSample.Create());
                        var workbench = window.ViewModel.InvoiceWorkbench;
                        workbench.InputPath = invoice;
                        workbench.SelectedOperation = workbench.Operations.Single(choice => choice.Value == OfficeIMO.Workflows.OfficeInvoiceWorkflowOperation.EditSource);
                        workbench.EditNumber = "EDITED-2026-0042";
                        if (scenario == "invoices-standards") {
                            var xml = System.Xml.Linq.XDocument.Load(invoice);
                            xml.Root!.SetAttributeValue("SchemaFailureProbe", "true"); xml.Save(invoice);
                            workbench.SelectedOperation = workbench.Operations.Single(choice => choice.Value == OfficeIMO.Workflows.OfficeInvoiceWorkflowOperation.Validate);
                            workbench.RequireStandards = true;
                            string rules = Path.GetFullPath(Path.Combine(root, "..", "invoicing-standards"));
                            workbench.RuleBundlePath = Path.Combine(rules, "xrechnung.zip");
                            workbench.SaxonJarPath = Path.Combine(rules, "saxon", "saxon-he-12.10.jar");
                        }
                        window.ViewModel.Commands["Invoices"].Execute(null);
                    }
                    if (scenario is "watermark" or "watermark-image") window.ViewModel.ShowEditModeCommand.Execute(null);
                    if (scenario == "connections") {
                        var connections = new ConnectionsWindow { DataContext = services.AiConnections };
                        _ = connections.ShowDialog<bool>(window);
                    }
                    Console.WriteLine($"READY {scenario} {width}x{height} {culture} {theme}");
                    Console.Out.Flush();
                } catch (Exception error) { Console.Error.WriteLine(error); }
            }, DispatcherPriority.Background)).StartWithClassicDesktopLifetime([]);
    }
}
