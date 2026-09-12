using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Assistant;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioConnectionExperienceTests {
    [Theory]
    [InlineData(4, "http://localhost:1234/v1/")]
    [InlineData(5, "http://localhost:11434/v1/")]
    public async Task LocalPresetsUseLocalTransportAndInvalidateEarlierConsent(int index, string endpoint) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var connections = TestAppBuilder.CreateTestServices().AiConnections;
            connections.ProviderIndex = 1;
            connections.ApiKey = "session-only";
            connections.Model = "old-model";
            connections.IsConnected = true;
            long revision = connections.Revision;
            connections.ProviderIndex = index;
            Assert.Equal(endpoint, connections.Options().Endpoint!.AbsoluteUri);
            Assert.True(connections.Profile().IsLocal);
            Assert.False(connections.CanUse);
            Assert.Empty(connections.ApiKey);
            Assert.Empty(connections.Model);
            Assert.True(connections.Revision > revision);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SavedChoicesRestoreWithoutCredentialsOrClaimingAConnection() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            var connections = services.AiConnections;
            connections.ProviderIndex = 4;
            connections.ApiKey = "never-write-this-key";
            connections.Model = "local-text-model";
            connections.IsConnected = true;
            Assert.True(connections.RememberSelection());
            string saved = File.ReadAllText(services.Paths.PreferencesPath);
            Assert.DoesNotContain("never-write-this-key", saved);
            var preferences = new StudioPreferencesService(new JsonStudioPreferencesStore(services.Paths.PreferencesPath));
            var restored = new StudioAiConnections(services.Paths.Root, services.Localizer, preferences);
            Assert.Equal(4, restored.ProviderIndex);
            Assert.Equal("local-text-model", restored.Model);
            Assert.Equal(connections.Endpoint, restored.Endpoint);
            Assert.Empty(restored.ApiKey);
            Assert.False(restored.IsConnected);
            Assert.False(restored.CanUse);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task LocalDiscoveryUsesTheRealAdapterWithoutSendingDocumentContent() {
        using var session = TestAppBuilder.StartSession();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        Task<string> request = ServeModelsAsync(listener, timeout.Token);
        await session.Dispatch(async () => {
            var connections = TestAppBuilder.CreateTestServices().AiConnections;
            connections.ProviderIndex = 4;
            connections.Endpoint = $"http://127.0.0.1:{port}/v1/";
            await connections.ConnectCommand.ExecuteAsync(null).WaitAsync(timeout.Token);
            Assert.True(connections.CanUse, connections.Status);
            Assert.Equal("local-text-model", connections.Model);
            Assert.Equal(["local-text-model"], connections.Models);
            return true;
        }, timeout.Token);
        string headers = await request;
        Assert.StartsWith("GET /v1/models HTTP/1.1", headers);
        Assert.DoesNotContain("Authorization:", headers, StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<string> ServeModelsAsync(TcpListener listener, CancellationToken token) {
        using TcpClient client = await listener.AcceptTcpClientAsync(token);
        await using NetworkStream stream = client.GetStream();
        using var reader = new StreamReader(stream, Encoding.ASCII, leaveOpen: true);
        var headers = new StringBuilder();
        while (await reader.ReadLineAsync(token) is { Length: > 0 } line) headers.AppendLine(line);
        byte[] body = Encoding.UTF8.GetBytes("""{"object":"list","data":[{"id":"local-text-model","object":"model","owned_by":"local"}]}""");
        byte[] response = Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: {body.Length}\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(response, token);
        await stream.WriteAsync(body, token);
        return headers.ToString();
    }

    [Theory]
    [InlineData(640, 540, 0, false, "en")]
    [InlineData(860, 760, 4, true, "en")]
    [InlineData(640, 540, 2, false, "en-XA")]
    public async Task SetupActionsRemainReachableAtSmallAndExpandedLayouts(int width, int height, int provider, bool light, string culture) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            IStudioLocalizer original = StudioLocalization.Current;
            var localizer = new StudioLocalizer(CultureInfo.GetCultureInfo(culture));
            StudioLocalization.Configure(localizer);
            var services = TestAppBuilder.CreateTestServices();
            var connections = new StudioAiConnections(services.Paths.Root, localizer) { ProviderIndex = provider };
            var window = new ConnectionsWindow { DataContext = connections, Width = width, Height = height,
                RequestedThemeVariant = light ? ThemeVariant.Light : ThemeVariant.Dark };
            try {
                window.Show(); window.UpdateLayout(); Dispatcher.UIThread.RunJobs();
                var footer = Assert.Single(window.GetVisualDescendants().OfType<Button>(),
                    button => Equals(button.Content, localizer.Get("Connections.UseConnection")));
                var position = footer.TranslatePoint(default, window)!.Value;
                Assert.True(footer.Bounds.Height >= 32);
                Assert.InRange(position.X, 0, width - footer.Bounds.Width + 1);
                Assert.InRange(position.Y, 0, height - footer.Bounds.Height + 1);
                Assert.False(footer.IsEffectivelyEnabled);
                Assert.Equal(connections.SelectedProvider.Name, connections.SelectedProvider.ToString());
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"connections-{provider}-{width}x{height}-{culture}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); StudioLocalization.Configure(original); }
            return true;
        }, CancellationToken.None);
    }
}
