using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.AI;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Tests;

public sealed partial class DocumentAssistantTests {
    [Theory]
    [InlineData(400, 540, false)]
    [InlineData(440, 620, true)]
    [InlineData(460, 820, false)]
    public async Task CitedAnswerExportsStayReachableAfterScrolling(int width, int height, bool light) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            using var model = new DocumentAssistantViewModel(connections, _ => new(bytes, "source.pdf", null, () => true),
                _ => { }, services.Localizer, (profile, _, _) => Task.FromResult<IOfficeAiExecutor>(new FixtureExecutor(profile, request => Task.FromResult(Answer(request))))) {
                CopyAnswer = _ => Task.CompletedTask, ExportAnswer = (_, _, _) => Task.FromResult(true)
            };
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            for (int i = 0; i < 3; i++) { model.Question = "What is in the source?"; await model.AskCommand.ExecuteAsync(null); }
            var view = new AssistantView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view, RequestedThemeVariant = light ? ThemeVariant.Light : ThemeVariant.Dark };
            try {
                window.Show(); window.UpdateLayout();
                view.FindControl<ScrollViewer>("ConversationScroll")!.ScrollToEnd();
                window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                foreach (var command in new System.Windows.Input.ICommand[] { model.CopyLastAnswerCommand, model.ExportLastAnswerCommand, model.AskCommand }) {
                    var button = Assert.Single(view.GetVisualDescendants().OfType<Button>(), item => ReferenceEquals(item.Command, command));
                    var position = button.TranslatePoint(default, window)!.Value;
                    Assert.True(button.Bounds.Height >= 30);
                    Assert.InRange(position.X, 0, width - button.Bounds.Width + 1);
                    Assert.InRange(position.Y, 0, height - button.Bounds.Height + 1);
                }
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (output is not null) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"answer-export-{width}x{height}-{(light ? "light" : "dark")}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
