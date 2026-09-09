using System.Text.Json;
using OfficeIMO.AI;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class DocumentAssistantTests {
    [Fact]
    public async Task AuthorizedEncryptedSourceAnswersWithoutSendingItsPassword() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            Directory.CreateDirectory(services.Paths.Root);
            string path = Path.Combine(services.Paths.Root, "protected.pdf");
            const string password = "private-reader-password";
            byte[] plain = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            File.WriteAllBytes(path, PdfDocument.Load(plain).Security.Encrypt(new(password) {
                OwnerPassword = "private-owner-password", AllowedPermissions = PdfStandardPermissions.CopyContents
            }).Pdf);
            using var workspace = await PdfWorkspace.OpenAsync(path, default, services.Recovery, password);
            Assert.True(workspace.ViewInfo.CanExtractContent);
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            var executor = new FixtureExecutor(connections.Profile(), request => {
                Assert.DoesNotContain(password, request.InputJson);
                Assert.DoesNotContain("private-owner-password", request.InputJson);
                return Task.FromResult(Answer(request));
            });
            using var model = new DocumentAssistantViewModel(connections,
                _ => new(workspace.CopyBytes(), workspace.FileName, null, () => true, workspace.CreateReaderOptions()),
                _ => { }, services.Localizer, (_, _, _) => Task.FromResult<IOfficeAiExecutor>(executor));
            model.Question = "What is in the source?";
            await model.AskCommand.ExecuteAsync(null);
            Assert.Single(model.Messages, message => !message.IsQuestion);
            Assert.Equal(File.ReadAllBytes(path), workspace.CopyBytes());
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task MultipleSavedAccountsRequireSelectionBeforeConnectionOrSignOut() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var store = new IntelligenceX.OpenAI.Auth.FileAuthBundleStore(Path.Combine(services.Paths.Root, "AI", "chatgpt-auth.json"));
            foreach (string id in new[] { "first", "second" }) await store.SaveAsync(new("openai-codex", "fixture", "fixture", null) { AccountId = id });
            await services.AiConnections.ConnectCommand.ExecuteAsync(null);
            Assert.False(services.AiConnections.IsConnected);
            Assert.Equal(2, services.AiConnections.Accounts.Count);
            await services.AiConnections.SignOutCommand.ExecuteAsync(null);
            Assert.Equal(2, (await store.ListAsync("openai-codex")).Count);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ChangedSourceDiscardsLateResultAndConversation() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var finish = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var executor = new FixtureExecutor(connections.Profile(), async request => {
                started.SetResult(); await finish.Task; return Answer(request);
            });
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            bool current = true;
            using var model = new DocumentAssistantViewModel(connections, _ => new(bytes, "source.pdf", null, () => current),
                _ => { }, services.Localizer, (_, _, _) => Task.FromResult<IOfficeAiExecutor>(executor));
            model.Question = "What is in the source?";
            Task pending = model.AskCommand.ExecuteAsync(null);
            await started.Task.WaitAsync(TimeSpan.FromSeconds(15));
            current = false; model.CheckSource(); finish.SetResult(); await pending;
            Assert.Empty(model.Messages);
            Assert.False(model.IsBusy);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ValidAnswerNavigatesToItsSourceAndConnectionChangeClearsIt() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            var executor = new FixtureExecutor(connections.Profile(), request => Task.FromResult(Answer(request)));
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            int navigated = 0;
            using var model = new DocumentAssistantViewModel(connections, _ => new(bytes, "source.pdf", null, () => true),
                page => navigated = page, services.Localizer, (_, _, _) => Task.FromResult<IOfficeAiExecutor>(executor));
            model.Question = "What is in the source?";
            await model.AskCommand.ExecuteAsync(null);
            var answer = Assert.Single(model.Messages, message => !message.IsQuestion);
            var citation = Assert.Single(answer.Citations);
            Assert.Contains("source sentence", citation.Quote);
            citation.NavigateCommand.Execute(null); Assert.Equal(1, navigated);
            connections.Model = "another-model";
            Assert.Empty(model.Messages);
            Assert.False(model.AllowRemoteProcessing);
            return true;
        }, CancellationToken.None);
    }

    private static OfficeAiExecutionResponse Answer(OfficeAiExecutionRequest request) {
        using var input = JsonDocument.Parse(request.InputJson);
        var evidence = input.RootElement.GetProperty("evidence")[0];
        string json = JsonSerializer.Serialize(new { claims = new[] { new {
            text = "The document contains a source sentence.", evidence = new[] { new { id = evidence.GetProperty("id").GetString(), quote = evidence.GetProperty("text").GetString() } }
        } }, fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>() });
        return new(json);
    }

    private sealed class FixtureExecutor(OfficeAiExecutionProfile profile, Func<OfficeAiExecutionRequest, Task<OfficeAiExecutionResponse>> run) : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile => profile;
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) => run(request);
    }
}
