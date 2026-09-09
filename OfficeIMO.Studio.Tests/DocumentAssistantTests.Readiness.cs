using OfficeIMO.AI;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Tests;

public sealed partial class DocumentAssistantTests {
    [Fact]
    public async Task CancelledProviderCannotPublishLateAnswerAndPreparedEvidenceCanBeReused() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var finish = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            using var model = new DocumentAssistantViewModel(connections, _ => new(bytes, "source.pdf", null, () => true),
                _ => { }, services.Localizer, (profile, _, _) => Task.FromResult<IOfficeAiExecutor>(new FixtureExecutor(profile, async request => {
                    started.TrySetResult(); await finish.Task; return Answer(request);
                })));
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            var prepared = model.PreparedDocument;
            model.Question = "Read the sentence";
            Task pending = model.AskCommand.ExecuteAsync(null);
            await started.Task.WaitAsync(TimeSpan.FromSeconds(15));
            model.CancelCommand.Execute(null);
            await pending.WaitAsync(TimeSpan.FromSeconds(5));
            finish.SetResult();
            Assert.Empty(model.LastAnswerText);
            Assert.DoesNotContain(model.Messages, message => !message.IsQuestion);
            Assert.Same(prepared, model.PreparedDocument);
            model.Question = "Try again";
            await model.AskCommand.ExecuteAsync(null);
            Assert.NotEmpty(model.LastAnswerText);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ExportRechecksSourceAfterDestinationSelection() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            bool current = true, checkedDestination = false;
            using var model = new DocumentAssistantViewModel(connections, _ => new(bytes, "source.pdf", null, () => current),
                _ => { }, services.Localizer, (profile, _, _) => Task.FromResult<IOfficeAiExecutor>(new FixtureExecutor(profile, request => Task.FromResult(Answer(request))))) {
                ExportAnswer = (_, isCurrent, _) => {
                    Assert.True(isCurrent()); current = false;
                    Assert.False(isCurrent()); checkedDestination = true;
                    return Task.FromResult(false);
                }
            };
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            model.Question = "What is in the source?";
            await model.AskCommand.ExecuteAsync(null);
            await model.ExportLastAnswerCommand.ExecuteAsync(null);
            Assert.True(checkedDestination);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task RepeatedQuestionsReuseEvidenceAndCopyTheReviewedAnswer() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            int captures = 0;
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            bool current = true;
            string? copied = null;
            using var model = new DocumentAssistantViewModel(connections, _ => {
                captures++; return new(bytes, "source.pdf", null, () => current);
            }, _ => { }, services.Localizer, (profile, _, _) => Task.FromResult<IOfficeAiExecutor>(new FixtureExecutor(profile, request => Task.FromResult(Answer(request))))) {
                CopyAnswer = text => { copied = text; return Task.CompletedTask; }
            };
            model.Question = "What is in the source?";
            Assert.False(model.CanAsk);
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            var prepared = model.PreparedDocument;
            Assert.NotNull(prepared);
            await model.AskCommand.ExecuteAsync(null);
            model.Question = "What sentence?";
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            await model.AskCommand.ExecuteAsync(null);
            Assert.Equal(1, captures);
            Assert.Same(prepared, model.PreparedDocument);
            await model.CopyLastAnswerCommand.ExecuteAsync(null);
            Assert.Contains(prepared.SourceHash, copied);
            Assert.Contains("A source sentence.", copied);
            current = false; model.CheckSource();
            Assert.False(model.CanReviewAnswer);
            Assert.Null(model.PreparedDocument);
            Assert.Empty(model.LastAnswerText);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task EmptyPageOffersOcrWithoutConnectingAndInactiveTabReleasesPreparedEvidence() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text(" ")))).ToBytes();
            bool openedOcr = false;
            using var model = new DocumentAssistantViewModel(services.AiConnections, _ => new(bytes, "scan.pdf", null, () => true),
                _ => { }, services.Localizer, (_, _, _) => throw new InvalidOperationException("Readiness must not connect")) {
                OpenOcr = _ => { openedOcr = true; return Task.CompletedTask; }
            };
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            Assert.NotNull(model.PreparedDocument);
            Assert.False(model.CanAsk);
            Assert.True(model.CanOpenOcr, model.EvidenceSummary + " / " + model.Status);
            await model.OpenOcrCommand.ExecuteAsync(null);
            Assert.True(openedOcr);
            model.Deactivate();
            Assert.Null(model.PreparedDocument);
            Assert.False(model.CanOpenOcr);
            return true;
        }, CancellationToken.None);
    }
}
