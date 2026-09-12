using Microsoft.JSInterop;

namespace OfficeIMO.Web.Converter.Services;

public sealed class ConverterInterop(IJSRuntime js) : IAsyncDisposable {
    internal const string ModulePath = "./Components/ConverterWorkspace.razor.js";
    internal const string CreateObjectUrlMethod = "createObjectUrl";
    internal const string RegisterWebMcpToolMethod = "registerWebMcpTool";
    internal const string UnregisterWebMcpToolMethod = "unregisterWebMcpTool";

    private IJSObjectReference? _module;

    private async ValueTask<IJSObjectReference> GetModuleAsync() =>
        _module ??= await js.InvokeAsync<IJSObjectReference>("import", ModulePath);

    public async ValueTask<string> CreateObjectUrlAsync(byte[] bytes, string contentType) {
        IJSObjectReference module = await GetModuleAsync();
        return await module.InvokeAsync<string>(CreateObjectUrlMethod, bytes, contentType);
    }

    public async ValueTask RevokeObjectUrlAsync(string? url) {
        if (string.IsNullOrWhiteSpace(url)) {
            return;
        }
        // Cleanup may finish after the component has disposed its imported module.
        await js.InvokeVoidAsync("URL.revokeObjectURL", url);
    }

    public async ValueTask RegisterWebMcpToolAsync<T>(DotNetObjectReference<T> converter) where T : class {
        IJSObjectReference module = await GetModuleAsync();
        await module.InvokeVoidAsync(RegisterWebMcpToolMethod, converter);
    }

    public async ValueTask UnregisterWebMcpToolAsync() {
        if (_module is null) {
            return;
        }
        try {
            await _module.InvokeVoidAsync(UnregisterWebMcpToolMethod);
        } catch (JSDisconnectedException) {
        }
    }

    public async ValueTask DisposeAsync() {
        try {
            if (_module is not null) {
                await _module.DisposeAsync();
            }
        } catch (JSDisconnectedException) {
        }
    }
}
