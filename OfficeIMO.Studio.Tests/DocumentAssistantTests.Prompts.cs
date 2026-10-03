using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.AI;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Tests;

public sealed partial class DocumentAssistantTests {
    [Theory]
    [InlineData(400, 540, false, false)]
    [InlineData(460, 820, true, true)]
    public async Task SummaryShortcutUsesSummaryAndEditedQuestionsUseAsk(int width, int height, bool light, bool edit) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 3; connections.Model = "fixture"; connections.IsConnected = true;
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("A source sentence.")))).ToBytes();
            string? operation = null;
            using var model = new DocumentAssistantViewModel(connections, _ => new(bytes, "source.pdf", null, () => true),
                _ => { }, services.Localizer, (profile, _, _) => Task.FromResult<IOfficeAiExecutor>(new FixtureExecutor(profile, request => {
                    using var input = JsonDocument.Parse(request.InputJson);
                    operation = input.RootElement.GetProperty("operation").GetString();
                    return Task.FromResult(Answer(request));
                })));
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            var view = new AssistantView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view, RequestedThemeVariant = light ? ThemeVariant.Light : ThemeVariant.Dark };
            try {
                window.Show(); window.UpdateLayout(); Dispatcher.UIThread.RunJobs();
                var summary = Assert.Single(view.GetVisualDescendants().OfType<Button>(), button => Equals(button.CommandParameter, "Summary"));
                var point = summary.TranslatePoint(new Point(summary.Bounds.Width / 2, summary.Bounds.Height / 2), window)!.Value;
                window.MouseDown(point, MouseButton.Left); window.MouseUp(point, MouseButton.Left);
                Assert.Equal(services.Localizer.Get("Assistant.Prompt.Summary"), model.Question);
                CapturePrompt("selected");
                if (edit) model.Question += " Focus on totals.";
                await model.AskCommand.ExecuteAsync(null);
                Assert.Equal(edit ? "Ask" : "Summarize", operation);
                Assert.Single(model.Messages, message => !message.IsQuestion);
                Assert.NotEmpty(Assert.Single(model.Messages, message => !message.IsQuestion).Citations);
                CapturePrompt("answer");

                // Sending clears the shortcut intent; a subsequent typed question stays Ask.
                model.Question = "What is in the source?";
                await model.AskCommand.ExecuteAsync(null);
                Assert.Equal("Ask", operation);

                void CapturePrompt(string state) {
                    window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                    using var frame = window.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (output is null) return;
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"summary-{width}x{height}-{(light ? "light" : "dark")}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
