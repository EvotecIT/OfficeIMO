using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Styling;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using System.Globalization;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFormOcrReviewTests {
    [Theory]
    [InlineData("en", 360, 760, false)]
    [InlineData("en", 680, 920, true)]
    [InlineData("pl", 360, 760, false)]
    [InlineData("de", 360, 760, true)]
    [InlineData("fr", 680, 920, true)]
    public async Task ReviewRequiresAcceptanceRejectsInvalidCorrectionsAndAppliesOneUndoableSaveCopy(string culture, int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var previousLocalizer = StudioLocalization.Current;
            CultureInfo previousCulture = CultureInfo.CurrentCulture, previousUi = CultureInfo.CurrentUICulture;
            CultureInfo? previousDefault = CultureInfo.DefaultThreadCurrentCulture, previousDefaultUi = CultureInfo.DefaultThreadCurrentUICulture;
            var paths = new StudioDataPaths(Path.Combine(((App)Application.Current!).Services.Paths.Root, culture));
            new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences { UiCulture = culture });
            var services = StudioApplicationServices.Create(paths);
            StudioLocalization.Configure(services.Localizer);
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "scan form.pdf"), copy = Path.Combine(services.Paths.Root, "filled copy.pdf");
            File.Copy(FixturePath(), source);
            byte[] original = File.ReadAllBytes(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(copy));
            await model.OpenDocumentAsync(source);
            await model.PrepareFormOcrAsync(Engine(PdfDocument.Load(original)));
            Assert.Null(model.FormOcrError);
            var review = Assert.IsType<FormOcrReviewViewModel>(model.FormOcrReview);
            await review.PreviewTask;
            Assert.False(model.CanApplyFormOcr);
            Assert.False(model.IsDirty);
            var amount = review.Proposals.Single(item => item.Proposal.Field.Name == "Amount");
            var country = review.Proposals.Single(item => item.Proposal.Field.Name == "Country");
            var code = review.Proposals.Single(item => item.Proposal.Field.Name == "SerialCode");
            var readonlyValue = review.Proposals.Single(item => item.Proposal.Field.Name == "ReadOnly");
            readonlyValue.Accepted = true; Assert.False(readonlyValue.Accepted);
            Assert.False(review.Proposals.Single(item => item.Proposal.Field.Name == "CustomRule").CanAccept);
            review.SelectedProposal = amount; await review.PreviewTask;
            var oldTheme = Application.Current!.RequestedThemeVariant;
            Application.Current.RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light;
            var view = new FormsInspectorView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = new ScrollViewer { Content = view } };
            try {
                window.Show(); await Render(window);
                view.GetVisualDescendants().OfType<Expander>().First().IsExpanded = true;
                await Render(window); Capture(window, $"form-ocr-review-{culture}-{width}");
                Assert.NotNull(review.PreviewRegion);
                review.FullPage = true; Assert.Null(review.PreviewRegion);
                review.FullPage = false; Assert.NotNull(review.PreviewRegion);
                var acceptance = view.GetVisualDescendants().OfType<CheckBox>().Single(box => Equals(box.Content, amount.AcceptValueLabel));
                await Click(window, acceptance); Assert.True(amount.Accepted);
                amount.Value = "1001.123";
                Assert.False(amount.Accepted);
                await Click(window, acceptance); Assert.True(amount.Accepted);
                Assert.False(model.CanApplyFormOcr); Assert.False(amount.IsValid); Assert.NotEmpty(amount.Validation);
                await model.ApplyFormOcrCommand.ExecuteAsync(null);
                Assert.False(model.IsDirty); Assert.Equal(original, File.ReadAllBytes(source));
                await Render(window); Capture(window, $"form-ocr-invalid-{culture}-{width}");
                amount.Value = "200.50"; Assert.False(amount.Accepted);
                await Click(window, acceptance); Assert.True(amount.Accepted);
                country.Value = "Germany"; country.Accepted = true;
                code.Value = "Z9Y8X7"; code.Accepted = true;
                Assert.True(model.CanApplyFormOcr);
                var manual = model.FormFields.Single(item => item.Name == "FullName");
                manual.TextValue = "Pending manual correction";
                Assert.True(model.HasFormDrafts);
                Assert.False(model.CanRecognizeFormValues); Assert.False(model.CanApplyFormOcr);
                await model.ApplyFormOcrCommand.ExecuteAsync(null);
                Assert.Equal(string.Empty, model.FormFields.Single(item => item.Name == "Amount").TextValue);
                manual.TextValue = string.Empty;
                Assert.True(model.CanRecognizeFormValues); Assert.True(model.CanApplyFormOcr);
                var apply = view.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, model.FormOcrApplyLabel));
                await Click(window, apply);
                await model.ApplyFormOcrCommand.ExecutionTask!;
                Assert.Null(model.ErrorMessage); Assert.True(model.IsDirty); Assert.True(model.CanUndo);
                Assert.Equal("200.50", model.FormFields.Single(item => item.Name == "Amount").TextValue);
                Assert.True(review.IsStale); Assert.False(model.CanApplyFormOcr);
                await Render(window); Capture(window, $"form-ocr-stale-{culture}-{width}");
                await model.UndoCommand.ExecuteAsync(null);
                Assert.Equal(string.Empty, model.FormFields.Single(item => item.Name == "Amount").TextValue);
                Assert.False(model.CanUndo); Assert.Equal(original, File.ReadAllBytes(source));
                await model.RedoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage); Assert.False(model.IsDirty);
                Assert.Equal("DE", PdfDocument.Load(copy).Inspect().FormFieldsByName["Country"].Value);
                Assert.Equal("Z9Y8X7", PdfDocument.Load(copy).Inspect().FormFieldsByName["SerialCode"].Value);
                Assert.Equal(original, File.ReadAllBytes(source));
                await Click(window, view.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, model.FormOcrCancelLabel)));
                Assert.False(model.HasFormOcrReview);
            } finally {
                window.Close(); Application.Current.RequestedThemeVariant = oldTheme;
                StudioLocalization.Configure(previousLocalizer);
                CultureInfo.CurrentCulture = previousCulture; CultureInfo.CurrentUICulture = previousUi;
                CultureInfo.DefaultThreadCurrentCulture = previousDefault; CultureInfo.DefaultThreadCurrentUICulture = previousDefaultUi;
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PreparationCapturesTheEngineOffDispatcherAndKeepsDispatcherLiveWhileProviderAwaits() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "cpu-prefix.pdf");
            File.Copy(FixturePath(), source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var engine = new ThreadObservedEngine(new DelegateOcrEngine("dispatcher-prefix", async (_, _) => {
                entered.TrySetResult(); await release.Task; return new OcrResult();
            }));
            Task preparation = model.PrepareFormOcrAsync(engine);
            try {
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(30));
                Assert.True(await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => true));
                Assert.False(preparation.IsCompleted);
                Assert.False(engine.CapturedOnUiThread);
            } finally { release.TrySetResult(); await preparation.WaitAsync(TimeSpan.FromSeconds(30)); }
            Assert.False(model.IsDirty);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SharedAdmissionAndActiveCloseWaitForRecognitionCancellationWithoutEditingEitherSource() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "form.pdf"), second = Path.Combine(services.Paths.Root, "second.pdf");
            File.Copy(FixturePath(), source); File.Copy(FixturePath(), second);
            byte[] original = File.ReadAllBytes(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            using var permitOne = await services.Jobs.EnterAsync(CancellationToken.None);
            using var permitTwo = await services.Jobs.EnterAsync(CancellationToken.None);
            int calls = 0;
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var cancellationSeen = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var slow = new ThreadObservedEngine(new DelegateOcrEngine("cancel-fixture", async (_, token) => {
                calls++; entered.TrySetResult();
                try { await Task.Delay(Timeout.Infinite, token); }
                finally { cancellationSeen.TrySetResult(); await release.Task; }
                return new OcrResult();
            }));
            Task recognition = model.PrepareFormOcrAsync(slow);
            Assert.True(model.CanCancelOperation); Assert.Equal(0, calls);
            model.CancelCurrentOperation(); await recognition.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.False(model.HasFormOcrReview); Assert.False(model.IsDirty);
            permitOne.Dispose(); permitTwo.Dispose();
            recognition = model.PrepareFormOcrAsync(slow);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(30));
            bool dispatched = await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => true);
            Assert.True(dispatched);
            Assert.False(await model.RequestCloseDocumentAsync());
            var owner = new Window();
            try {
                owner.Show();
                var dialog = new ActiveOperationsDialog([model], services.Localizer);
                Task<bool> closing = dialog.ShowDialog<bool>(owner);
                dialog.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "CancelWorkAndClose")
                    .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                await cancellationSeen.Task.WaitAsync(TimeSpan.FromSeconds(30));
                await Task.Delay(100);
                Assert.True(model.CanCancelOperation); Assert.False(closing.IsCompleted);
                release.TrySetResult(); await recognition.WaitAsync(TimeSpan.FromSeconds(30));
                Assert.True(await closing.WaitAsync(TimeSpan.FromSeconds(10)));
                Assert.False(model.CanCancelOperation);
            } finally { release.TrySetResult(); model.CancelCurrentOperation(); await recognition; owner.Close(); }
            Assert.False(slow.CapturedOnUiThread);
            await model.OpenDocumentAsync(second);
            Assert.False(model.HasFormOcrReview); Assert.False(model.IsDirty);
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.Equal(original, File.ReadAllBytes(second));
            return true;
        }, CancellationToken.None);
    }

    // Engine identity/capabilities are read by the canonical runner after snapshot preparation,
    // before its first provider await. Provider callbacks themselves already run on a worker.
    private sealed class ThreadObservedEngine(IOcrEngine engine) : IOcrEngine {
        private int _capturedThread;
        public bool CapturedOnUiThread => Volatile.Read(ref _capturedThread) == 1;
        public string Id { get { Capture(); return engine.Id; } }
        public OcrEngineCapabilities Capabilities { get { Capture(); return engine.Capabilities; } }
        public Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken token = default) => engine.RecognizeAsync(request, token);
        private void Capture() => Interlocked.CompareExchange(ref _capturedThread, Avalonia.Threading.Dispatcher.UIThread.CheckAccess() ? 1 : 2, 0);
    }

    [Fact]
    public async Task AmbiguousWidgetsKeepRepeatedWordEvidenceAndExposeEachCapturedRegion() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes("%PDF-1.7\n" +
                "1 0 obj << /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >> endobj\n" +
                "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n" +
                "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [7 0 R 8 0 R] >> endobj\n" +
                "5 0 obj << /Fields [6 0 R] /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 9 0 R >> >> >> endobj\n" +
                "6 0 obj << /FT /Tx /T (Destination) /V () /Kids [7 0 R 8 0 R] >> endobj\n" +
                "7 0 obj << /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 200 200 230] /P 3 0 R >> endobj\n" +
                "8 0 obj << /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 100 200 130] /P 3 0 R >> endobj\n" +
                "9 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n" +
                "trailer << /Root 1 0 R /Size 10 >>\n%%EOF\n");
            var engine = new DelegateOcrEngine("multi-widget-fixture", (_, _) => Task.FromResult(new OcrResult {
                Spans = [Span("Bora", 30, 75), Span("Bora", 65, 75), Span("Moorea", 30, 175)]
            }));
            var prepared = await OfficeIMO.Pdf.Ocr.PdfFormOcrExtensions.PrepareFormOcrAsync(PdfDocument.Load(bytes), engine);
            using var review = new FormOcrReviewViewModel(prepared, ((App)Application.Current!).Services.Localizer);
            var proposal = Assert.Single(review.Proposals);
            Assert.True(proposal.Proposal.IsAmbiguous);
            Assert.Equal("Bora Bora", proposal.OriginalText);
            Assert.True(review.HasMultipleSources); Assert.Equal(2, review.Sources.Count);
            await review.PreviewTask;
            var firstRegion = review.PreviewRegion;
            review.SelectedSource = review.Sources[1]; await review.PreviewTask;
            Assert.NotEqual(firstRegion, review.PreviewRegion); Assert.NotNull(review.Preview);
            Assert.False(review.CanApply);
            proposal.Value = "Reviewed destination"; proposal.Accepted = true;
            Assert.True(review.CanApply);
            return true;
        }, CancellationToken.None);

        static OcrTextSpan Span(string text, double x, double y) => new() {
            Text = text, Confidence = .99, Level = OcrTextSpanLevel.Word,
            CoordinateUnit = OcrCoordinateUnit.Points, Region = new() { X = x, Y = y, Width = 30, Height = 14 }
        };
    }

    private static string FixturePath() => Path.Combine(AppContext.BaseDirectory, "Fixtures", "PdfFormOcr", "reportlab-scanned-form.pdf");
    private static IOcrEngine Engine(PdfDocument source) {
        var fields = source.Inspect().FormFields;
        var layouts = source.GetPageLayouts().ToDictionary(page => page.PageNumber);
        var text = new Dictionary<string, string> { ["FullName"] = "Alex Morgan", ["SerialCode"] = "A1B2C3", ["Country"] = "Poland",
            ["Amount"] = "123.45", ["Reference"] = "INV-2048", ["ReadOnly"] = "KEEP", ["CustomRule"] = "77" };
        return new DelegateOcrEngine("fixture", (request, _) => Task.FromResult(new OcrResult {
            Spans = fields.SelectMany(field => field.Widgets.Where(widget => widget.PageNumber == request.PageNumber).Select(widget => {
                var bounds = layouts[widget.PageNumber!.Value].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
                return new OcrTextSpan { Text = text[field.Name!], Confidence = field.Name == "SerialCode" ? 0.1 : 0.99,
                    Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
                    Region = new() { X = bounds.Left + 2, Y = bounds.Top + 2, Width = bounds.Width - 4, Height = bounds.Height - 4 } };
            })).ToArray()
        }));
    }
    private static async Task Render(Window window) {
        await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(window.UpdateLayout, Avalonia.Threading.DispatcherPriority.Background);
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }
    private static async Task Click(Window window, Control control) {
        await Render(window);
        control.BringIntoView(); await Render(window);
        Point point = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width); Assert.InRange(point.Y, 0, window.Bounds.Height);
        window.MouseDown(point, Avalonia.Input.MouseButton.Left); window.MouseUp(point, Avalonia.Input.MouseButton.Left);
        await Render(window);
    }
    private static void Capture(Window window, string name) {
        using var bitmap = window.CaptureRenderedFrame(); Assert.NotNull(bitmap);
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(path)) return;
        Directory.CreateDirectory(path); bitmap.Save(Path.Combine(path, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
