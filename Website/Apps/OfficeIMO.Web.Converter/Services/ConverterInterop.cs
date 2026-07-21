using Microsoft.JSInterop;

namespace OfficeIMO.Web.Converter.Services;

public sealed class ConverterInterop(IJSRuntime js) : IAsyncDisposable {
    internal const string ModulePath = "./Components/ConverterWorkspace.razor.js";
    internal const string CreateObjectUrlMethod = "createObjectUrl";
    internal const string RevokeObjectUrlMethod = "revokeObjectUrl";

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
        IJSObjectReference module = await GetModuleAsync();
        await module.InvokeVoidAsync(RevokeObjectUrlMethod, url);
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
