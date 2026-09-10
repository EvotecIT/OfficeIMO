using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Tests;

public sealed partial class DocumentAssistantTests {
    [Theory]
    [InlineData("source-check")]
    [InlineData("source-prepare")]
    [InlineData("scope")]
    public async Task ChangedEvidenceRequiresFreshRemoteConsent(string change) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 1; connections.Model = "fixture"; connections.IsConnected = true;
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Source evidence.")))).ToBytes();
            int revision = 0;
            using var model = new DocumentAssistantViewModel(connections, currentPage => {
                int captured = revision;
                return new(bytes, "source.pdf", currentPage ? 1 : null, () => captured == revision);
            }, _ => { }, services.Localizer);
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            model.Question = "What is in the source?";
            model.AllowRemoteProcessing = true;
            Assert.True(model.CanAsk);
            model.NewConversationCommand.Execute(null);
            Assert.True(model.AllowRemoteProcessing);
            Assert.True(model.CanAsk);
            if (change == "scope") {
                model.CurrentPageOnly = true;
                await model.PrepareEvidenceCommand.ExecutionTask!;
            } else {
                revision++;
                if (change == "source-check") model.CheckSource();
                await model.PrepareEvidenceCommand.ExecuteAsync(null);
            }
            Assert.NotNull(model.PreparedDocument);
            Assert.False(model.AllowRemoteProcessing);
            Assert.False(model.CanAsk);
            model.AllowRemoteProcessing = true;
            Assert.True(model.CanAsk);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task RefreshingSavedAccountsPreservesTheRenderedSelectionAndConnectionRevision() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.RefreshAccounts(["first", "selected"]);
            connections.AccountId = "selected";
            var view = new ConnectionsView { DataContext = connections };
            var window = new Window { Content = view, Width = 480, Height = 720 };
            try {
                window.Show();
                window.UpdateLayout(); Dispatcher.UIThread.RunJobs();
                var selector = Assert.Single(view.GetVisualDescendants().OfType<ComboBox>(), box => ReferenceEquals(box.ItemsSource, connections.Accounts));
                Assert.Equal("selected", selector.SelectedItem);
                long revision = connections.Revision;
                connections.RefreshAccounts(["selected", "new-account"]);
                window.UpdateLayout(); Dispatcher.UIThread.RunJobs();
                Assert.Equal("selected", selector.SelectedItem);
                Assert.Equal("selected", connections.AccountId);
                Assert.Equal(revision, connections.Revision);
                Assert.Equal(["selected", "new-account"], connections.Accounts);
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, "assistant-account-refresh-480x720.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                selector.SelectedItem = "new-account";
                Assert.Equal("new-account", connections.AccountId);
                Assert.True(connections.Revision > revision);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
