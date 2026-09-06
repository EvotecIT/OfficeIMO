using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSigningPreviewTests {
    [Fact]
    public async Task ApplyingAFormDraftBeforeReviewSignsTheChosenValue() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services; Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "form.pdf"), output = Path.Combine(services.Paths.Root, "signed.pdf");
            PdfDocument.Create(document => document.Page(page => page.Content(content => content.Item(item => item.TextField("Name", value: "Original"))))).Save(source);
            using var certificate = CreateCertificate();
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output), reviewSigning: _ => Task.FromResult(true),
                loadSigningCertificate: _ => new X509Certificate2(certificate));
            await model.OpenDocumentAsync(source); SelectCertificate(model, certificate); model.SignatureIsVisible = false;
            model.FormFields[0].TextValue = "Reviewed value"; await model.ApplyFormDraftsCommand.ExecuteAsync(null);
            Assert.False(model.HasFormDrafts); Assert.True(model.IsDirty);
            await model.ApplyCertificateSignatureCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage); Assert.True(model.IsDirty);
            Assert.Equal("Reviewed value", PdfDocument.Load(output).Inspect().FormFieldsByName["Name"].Value);
            Assert.Equal("Original", PdfDocument.Load(source).Inspect().FormFieldsByName["Name"].Value);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DeclinedReviewOrChangedCertificateNeverProducesASignedCopy(bool changedCertificate) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services; Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "signed.pdf");
            CreatePdf().Save(source); byte[] original = File.ReadAllBytes(source);
            using var selected = CreateCertificate(); using var different = CreateCertificate(); int loads = 0;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output), reviewSigning: _ => Task.FromResult(changedCertificate),
                loadSigningCertificate: _ => { loads++; return new X509Certificate2(different); });
            await model.OpenDocumentAsync(source); SelectCertificate(model, selected);
            await model.ApplyCertificateSignatureCommand.ExecuteAsync(null);
            Assert.Equal(changedCertificate ? 1 : 0, loads); Assert.False(File.Exists(output));
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.False(model.IsDirty);
            if (changedCertificate) { Assert.NotNull(model.ErrorMessage); Assert.False(Assert.Single(services.Jobs.Entries).IsActive); }
            else Assert.Empty(services.Jobs.Entries);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CancelledSigningAtTheCpuGateLeavesNoOutputOrRecovery() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services; Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "signed.pdf"); CreatePdf().Save(source);
            using var certificate = CreateCertificate();
            using var blocker = await Features.Workspace.PdfWorkspace.OpenAsync(source, default);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously); Task<bool>? holding = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output), reviewSigning: async _ => {
                    holding = blocker.RunNonDetachableCpuWorkAsync(() => { acquired.SetResult(); release.Task.GetAwaiter().GetResult(); return true; }, default);
                    await acquired.Task; return true;
                }, loadSigningCertificate: _ => new X509Certificate2(certificate));
            try {
                await model.OpenDocumentAsync(source); SelectCertificate(model, certificate);
                Task pending = model.ApplyCertificateSignatureCommand.ExecuteAsync(null);
                DateTime deadline = DateTime.UtcNow.AddSeconds(10);
                while (services.Jobs.Entries.Count == 0 && DateTime.UtcNow < deadline) await Task.Delay(10);
                var job = Assert.Single(services.Jobs.Entries); job.CancelCommand.Execute(null); await pending;
                Assert.False(job.IsActive); Assert.False(job.HasOutput); Assert.False(job.HasRecovery);
                Assert.False(File.Exists(output)); Assert.False(model.IsDirty);
            } finally { release.TrySetResult(); if (holding is not null) await holding; }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CapturedSettingsSignUnsavedContentAndRejectStaleApproval(bool stale) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services; Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "signed.pdf");
            CreatePdf().Save(source); byte[] original = File.ReadAllBytes(source);
            using var certificate = CreateCertificate(); int loads = 0;
            MainWindowViewModel? model = null; PdfSigningPreviewViewModel? result = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => { model!.SignatureFieldName = "Later"; model.SignaturePageNumber = 100; return Task.FromResult<string?>(output); },
                reviewSigning: async preview => { Assert.Equal(0, loads); Assert.NotNull(preview.PreviewImage); if (stale) await model!.RotateRightCommand.ExecuteAsync(null); return true; },
                showSigningResult: preview => { result = preview; return Task.CompletedTask; },
                loadSigningCertificate: _ => { loads++; return new X509Certificate2(certificate); })) {
                await model.OpenDocumentAsync(source); SelectCertificate(model, certificate);
                model.SetOrganizerSelection([model.OrganizerPages[0]]); await model.DuplicateSelectedCommand.ExecuteAsync(null);
                await model.ApplyCertificateSignatureCommand.ExecuteAsync(null);
                Assert.Equal(original, File.ReadAllBytes(source)); Assert.True(model.IsDirty); Assert.Equal(3, model.Pages.Count);
                Assert.Equal("Later", model.SignatureFieldName);
                if (stale) { Assert.Equal(0, loads); Assert.Empty(services.Jobs.Entries); Assert.Null(result); Assert.False(File.Exists(output)); }
                else {
                    Assert.Null(model.ErrorMessage); Assert.Equal(1, loads); Assert.NotNull(result); Assert.True(result.CanOpenOutput);
                    var signed = PdfDocument.Load(File.ReadAllBytes(output)); Assert.Equal(3, signed.Inspect().PageCount);
                    var report = signed.Security.ValidateSignatures(new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default));
                    Assert.Equal("Approval", Assert.Single(report.Signatures).Signature.FieldName); Assert.True(report.MathematicalSignaturesVerified);
                    Assert.Contains("Certificate chain: not verified", result.Verification);
                    Assert.False(Assert.Single(services.Jobs.Entries).IsActive);
                    await model.UndoCommand.ExecuteAsync(null); Assert.Equal(2, model.Pages.Count); Assert.False(model.CanUndo);
                }
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ProviderSigningRequiresConsentAndKeepsVerifiedRecovery(bool consent, bool fail) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services; Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"); CreatePdf().Save(source);
            var output = new TestStorageFile("content://signing/output", []) { FailWrite = fail };
            string destination = await services.Storage.RegisterAsync(output.Item, default);
            using var certificate = CreateCertificate(); int loads = 0; PdfSigningPreviewViewModel? result = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(destination), reviewSigning: _ => Task.FromResult(true),
                confirmProviderWrite: _ => Task.FromResult(consent), showSigningResult: async preview => {
                    result = preview; var dialog = new PdfSigningDialog(preview) { Width = 380, Height = 440 };
                    try { dialog.Show(); await Layout(dialog); Capture(dialog, $"signing-provider-{fail}"); } finally { dialog.Close(); }
                },
                loadSigningCertificate: _ => { loads++; return new X509Certificate2(certificate); });
            await model.OpenDocumentAsync(source); SelectCertificate(model, certificate); model.SignatureIsVisible = false;
            await model.ApplyCertificateSignatureCommand.ExecuteAsync(null);
            Assert.False(model.IsDirty);
            if (!consent) { Assert.Equal(0, loads); Assert.Equal(0, output.Writes); Assert.Null(result); Assert.Empty(services.Jobs.Entries); }
            else {
                Assert.NotNull(result); Assert.Equal(!fail, result.CanOpenOutput); Assert.Equal(fail, result.HasRecovery);
                Assert.False(result.CanRevealOutput); Assert.False(Assert.Single(services.Jobs.Entries).IsActive);
                byte[] bytes = output.Bytes;
                if (fail) { var recovery = Assert.Single(services.WorkflowRecovery.GetRecoveries()); await services.WorkflowRecovery.VerifyAsync(recovery); bytes = File.ReadAllBytes(recovery.FilePath); Assert.Contains("destination contents are unconfirmed", result.Verification); }
                Assert.True(PdfDocument.Load(bytes).Security.ValidateSignatures(new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default)).MathematicalSignaturesVerified);
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(380, 440, false)]
    [InlineData(620, 700, true)]
    public async Task RenderedSigningReviewShowsAppearanceAndOpensTheVerifiedCopy(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root); string source = Path.Combine(services.Paths.Root, "source.pdf"); CreatePdf().Save(source);
            string output = Path.Combine(services.Paths.Root, new string('s', 120) + ".pdf");
            using var certificate = CreateCertificate();
            var owner = new MainWindow(services) { Width = 960, Height = 620 };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output),
                reviewSigning: preview => new PdfSigningDialog(preview) { Width = width, Height = height }.ShowDialog<bool>(owner),
                showSigningResult: preview => new PdfSigningDialog(preview) { Width = width, Height = height }.ShowDialog(owner),
                openDocumentInTab: (path, token) => owner.TabHost.OpenDocumentAsync(path, token),
                loadSigningCertificate: _ => new X509Certificate2(certificate));
            try {
                owner.Show(); await model.OpenDocumentAsync(source); SelectCertificate(model, certificate);
                Task signing = model.ApplyCertificateSignatureCommand.ExecuteAsync(null);
                var dialog = await WaitDialog(owner, false); await Layout(dialog); Capture(dialog, $"signing-preview-{width}-{dark}");
                Assert.NotNull(((PdfSigningPreviewViewModel)dialog.DataContext!).PreviewImage);
                dialog.FindControl<ScrollViewer>("ReviewContent")!.ScrollToEnd(); await Layout(dialog); Capture(dialog, $"signing-review-end-{width}-{dark}");
                Click(dialog, services.Localizer.Get("Signing.Create"));
                var resultDialog = await WaitDialog(owner, true); var result = (PdfSigningPreviewViewModel)resultDialog.DataContext!;
                await Layout(resultDialog); Capture(resultDialog, $"signing-result-{width}-{dark}");
                Assert.Contains("Signature math and document digest: verified", result.Verification);
                Assert.Contains("Certificate chain: not verified", result.Verification);
                Click(resultDialog, services.Localizer.Get("Organizer.ExtractOpen")); await result.OpenOutputCommand.ExecutionTask!;
                Assert.True(owner.ViewModel.HasDocumentSignatures); Assert.Equal(2, owner.ViewModel.Pages.Count);
                Click(resultDialog, services.Localizer.Get("Common.Close")); await signing;
                Assert.Null(model.ErrorMessage); Assert.False(model.IsDirty);
            } finally { foreach (var dialog in owner.OwnedWindows.ToArray()) dialog.Close(); owner.Close(); owner.TabHost.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    private static async Task<PdfSigningDialog> WaitDialog(Window owner, bool result) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(15);
        while (DateTime.UtcNow < deadline) {
            var dialog = owner.OwnedWindows.OfType<PdfSigningDialog>().SingleOrDefault();
            if (dialog?.DataContext is PdfSigningPreviewViewModel model && model.HasResult == result) return dialog;
            await Task.Delay(10);
        }
        throw new TimeoutException("Signing dialog did not open.");
    }
    private static async Task Layout(Window window) { window.UpdateLayout(); await Task.Delay(30); window.UpdateLayout(); }
    private static void Click(Window window, string label) {
        var button = window.GetVisualDescendants().OfType<Button>().Single(item => Equals(item.Content, label));
        Assert.True(button.IsVisible); Assert.True(button.IsEnabled); var point = button.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.Y + button.Bounds.Height, 0, window.Bounds.Height);
        button.Focus(); button.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
    }
    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_SIGNING_CAPTURE_DIR"); if (string.IsNullOrEmpty(root)) return;
        Directory.CreateDirectory(root); using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
    private static PdfDocument CreatePdf() => PdfDocument.Create(document => {
        document.Page(page => page.Size(300, 400).Content(content => content.Text("Document for signing review")));
        document.Page(page => page.Size(300, 400));
    });
    private static void SelectCertificate(MainWindowViewModel model, X509Certificate2 certificate) {
        model.SelectedSigningCertificate = new(certificate.Thumbprint, "Studio Test Signer", certificate.NotAfter, certificate.Issuer);
        model.SignatureFieldName = "Approval"; model.SignatureWidth = 200; model.SignatureHeight = 50;
    }
    private static X509Certificate2 CreateCertificate() {
        using var rsa = RSA.Create(2048); var request = new CertificateRequest("CN=Studio Test Signer", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }
}
