using System.Reflection;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using OfficeIMO.Web.Converter.Components;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class TextIntegritySelectionLifetimeTests {
    [Fact]
    public async Task EditingSourceInvalidatesCopiesReportsAndReviewBeforeCleanupCompletes() {
        var js = new DelayedRevocation();
        var component = new TextIntegrityWorkbench();
        Set(component, "_interop", new ConverterInterop(js));
        Set(component, "_outputUrl", "blob:copy"); Set(component, "_reportUrl", "blob:report");
        Set(component, "_text", "old"); Set(component, "_cleaned", "old copy");
        var pending = Invoke(component, "TextChangedAsync", new ChangeEventArgs { Value = "new source" });
        await js.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Null(Get(component, "_outputUrl")); Assert.Null(Get(component, "_reportUrl")); Assert.Null(Get(component, "_review")); Assert.Null(Get(component, "_cleaned"));
        Assert.Equal("new source", Get(component, "_text"));
        Assert.DoesNotContain("ready", (string)Get(component, "_message")!);
        js.Release.SetResult(); await pending.WaitAsync(TimeSpan.FromSeconds(5));
        await component.DisposeAsync(); Assert.Equal(2, js.Revoked.Count);
    }
    [Fact]
    public async Task DisposedInspectionCannotPublishAReviewAfterPriorUrlCleanupResumes() {
        var js = new DelayedRevocation();
        var component = new TextIntegrityWorkbench(); Set(component, "_interop", new ConverterInterop(js));
        Set(component, "_text", "review\u202Ethis"); Set(component, "_reportUrl", "blob:old");
        var pending = Invoke(component, "InspectAsync");
        await js.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await component.DisposeAsync(); js.Release.SetResult(); await pending.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Null(Get(component, "_review")); Assert.Null(Get(component, "_reportUrl"));
    }
    private static Task Invoke(object target, string name, params object[] args) => (Task)target.GetType().GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(target, args)!;
    private static object? Get(object target, string name) => target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(target);
    private static void Set(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(target, value);
    private sealed class DelayedRevocation : IJSRuntime {
        internal TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal List<string> Revoked { get; } = [];
        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args) => InvokeAsync<TValue>(identifier, default, args);
        public async ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken token, object?[]? args) {
            Assert.Equal("URL.revokeObjectURL", identifier); Revoked.Add((string)args![0]!);
            Entered.TrySetResult(); await Release.Task; return default!;
        }
    }
}
