using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioReviewedCopyDraftTests {
    [Theory]
    [InlineData("sign", false)]
    [InlineData("sign", true)]
    [InlineData("protect", false)]
    [InlineData("protect", true)]
    [InlineData("extract", false)]
    [InlineData("extract", true)]
    [InlineData("split", false)]
    [InlineData("split", true)]
    public async Task ReviewedCopiesRejectDraftsBeforeAndDuringReview(string operation, bool duringReview) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services; Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "form.pdf"), output = Path.Combine(services.Paths.Root, "copy.pdf");
            byte[] form = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Item(item => item.TextField("Name", value: "Original"))))).ToBytes();
            bool retained = operation is "extract" or "split";
            File.WriteAllBytes(source, retained ? PdfDocument.Create(document => document.Page(page => page.Size(300, 400))).ToBytes() : form);
            byte[] original = File.ReadAllBytes(source); MainWindowViewModel? model = null; int reviews = 0, pickers = 0;
            void AddDraft() {
                if (retained) model!.UnassignedFormDrafts.Add(new PdfFormFieldViewModel(PdfDocument.Load(form).Inspect().FormFields[0]) { TextValue = "Retained draft" });
                else model!.FormFields[0].TextValue = "Unapplied value";
            }
            Task<bool> Review() { reviews++; AddDraft(); return Task.FromResult(true); }
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => { pickers++; return Task.FromResult<string?>(output); },
                pickOutputFolder: _ => { pickers++; return Task.FromResult<string?>(Path.Combine(services.Paths.Root, "parts")); },
                reviewSigning: _ => Review(), reviewProtection: _ => Review(), reviewPageExtraction: _ => Review(), reviewPageSplit: _ => Review(),
                loadSigningCertificate: _ => throw new InvalidOperationException("Private key must not be loaded for unresolved drafts."))) {
                await model.OpenDocumentAsync(source); model.SetOrganizerSelection(model.OrganizerPages);
                model.SelectedSigningCertificate = new("test", "Test signer", DateTime.Now.AddDays(1), "Test issuer");
                model.SignatureIsVisible = false;
                model.ProtectUserPassword = "reader"; model.ProtectConfirmPassword = "reader";
                if (!duringReview) AddDraft();
                await (operation switch {
                    "sign" => model.ApplyCertificateSignatureCommand.ExecuteAsync(null),
                    "protect" => model.SaveProtectedCopyCommand.ExecuteAsync(null),
                    "extract" => model.ExtractSelectedCommand.ExecuteAsync(null),
                    _ => model.SplitCommand.ExecuteAsync(null)
                });
                Assert.Equal(duringReview ? 1 : 0, reviews); Assert.Equal(duringReview ? 1 : 0, pickers);
                Assert.True(model.HasFormDrafts); Assert.True(model.IsDirty);
                Assert.Equal(services.Localizer.Get("Workspace.CopyHasFormDrafts"), model.ErrorMessage);
                Assert.Empty(services.Jobs.Entries); Assert.False(File.Exists(output));
                Assert.Equal(original, File.ReadAllBytes(source)); Assert.False(model.CanUndo);
            }
            return true;
        }, CancellationToken.None);
    }
}
