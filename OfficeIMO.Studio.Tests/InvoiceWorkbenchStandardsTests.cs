using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioInvoiceStandardsTheoryAttribute : TheoryAttribute {
    public StudioInvoiceStandardsTheoryAttribute() {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_STANDARDS_TESTS") != "1")
            Skip = "Requires explicitly configured pinned invoice artifacts and Saxon runtime files.";
    }
}

public sealed class InvoiceWorkbenchStandardsTests {
    [StudioInvoiceStandardsTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RequestedButBlockedStagesKeepCapturedIntentAndProduceNoOutput(bool mismatch) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string input = Path.Combine(services.Paths.Root, "schema-probe.xml");
            XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(InvoiceSample.Create()));
            if (!mismatch) xml.Root!.SetAttributeValue("InvalidSchemaProbe", "true");
            xml.Save(input);
            InvoiceWorkbenchViewModel? model = null;
            using var owned = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null),
                new CapturedRunner(() => { if (mismatch) model!.RequireStandards = false; })) {
                InputPath = input, RequireStandards = true,
                RuleBundlePath = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE")!,
                SaxonJarPath = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!
            };
            model = owned;
            model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.Validate);
            if (mismatch) model.SelectedRelease = model.Releases.Single(r => r.Value == InvoiceSpecificationRelease.PeppolBis_3_0_21);
            var view = new InvoiceWorkbenchView { DataContext = model };
            var window = new Window { Width = 960, Height = 620, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                var settings = view.GetVisualDescendants().OfType<Expander>().Single(e => e.IsEffectivelyVisible);
                settings.IsExpanded = true; window.UpdateLayout();
                ((ScrollViewer)view.Content!).Offset = new Vector(0, 200); window.UpdateLayout();
                Capture(window, "invoice-standards-settings.png");
                await model.RunCommand.ExecuteAsync(null);
                Assert.False(model.HasOutput);
                Assert.Equal(mismatch ? "Not run" : "Invalid", model.SchemaSummary);
                Assert.Equal("Not run", model.RulesSummary);
                Assert.Contains(model.Diagnostics, d => d.Contains(mismatch ? "INV-RULESET-PROFILE" : "XSD", StringComparison.Ordinal));
                ((ScrollViewer)view.Content!).Offset = new Vector(0, double.MaxValue); window.UpdateLayout();
                Capture(window, mismatch ? "invoice-standards-mismatch.png" : "invoice-standards-schema-failure.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private sealed class CapturedRunner(Action entered) : IOfficeInvoiceWorkflowRunner {
        public Task<OfficeInvoiceStorageWorkflowResult> RunInvoiceAsync(OfficeInvoiceStorageWorkflowRequest request,
            InvoiceValidator? validator = null, CancellationToken cancellationToken = default) {
            entered(); return new OfficeWorkflowRunner().RunInvoiceAsync(request, validator, cancellationToken);
        }
    }
    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
