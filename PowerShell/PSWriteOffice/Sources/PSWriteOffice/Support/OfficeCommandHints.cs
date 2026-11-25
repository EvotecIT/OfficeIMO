using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace PSWriteOffice.Support;

/// <summary>Warns when a retired PSWriteOffice command is called and names its replacement.</summary>
/// <remarks>Each runspace that imports the module chains one lookup handler after any earlier one. The hint table is
/// read on the first lookup that misses, so importing the module costs only the registration.</remarks>
internal static class OfficeCommandHints {
    private static readonly ConditionalWeakTable<Runspace, Registration> Registrations = new();
    private static readonly Lazy<Dictionary<string, string>> Hints = new(LoadHints);

    internal static void Register() {
        Runspace? runspace = Runspace.DefaultRunspace;
        if (runspace == null) return;
        lock (Registrations) {
            // A repeated import replaces this module's handler instead of chaining a second copy.
            if (Registrations.TryGetValue(runspace, out Registration? existing)) {
                existing.Unregister();
                Registrations.Remove(runspace);
            }
            // Import runs on the pipeline thread, where SessionStateProxy refuses calls while the import pipeline runs.
            if (ScriptBlock.Create("$ExecutionContext").Invoke().FirstOrDefault()?.BaseObject is not EngineIntrinsics engine) return;
            Registrations.Add(runspace, new Registration(engine.InvokeCommand));
        }
    }

    internal static void Unregister() {
        Runspace? runspace = Runspace.DefaultRunspace;
        if (runspace == null) return;
        lock (Registrations) {
            if (Registrations.TryGetValue(runspace, out Registration? registration)) {
                registration.Unregister();
                Registrations.Remove(runspace);
            }
        }
    }

    private static Dictionary<string, string> LoadHints() {
        using Stream stream = typeof(OfficeCommandHints).Assembly.GetManifestResourceStream("PSWriteOffice.CommandHints.json")
            ?? throw new InvalidOperationException("The PSWriteOffice command-hint resource is missing. Rebuild the module.");
        using JsonDocument document = JsonDocument.Parse(stream);
        var hints = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (JsonProperty property in document.RootElement.EnumerateObject()) {
            if (property.Value.ValueKind == JsonValueKind.String) hints[property.Name] = property.Value.GetString()!;
        }
        return hints;
    }

    private sealed class Registration {
        private readonly CommandInvocationIntrinsics _commands;
        private readonly EventHandler<CommandLookupEventArgs>? _previous;
        private readonly EventHandler<CommandLookupEventArgs> _handler;
        private ScriptBlock? _warning;
        private volatile bool _active = true;

        internal Registration(CommandInvocationIntrinsics commands) {
            _commands = commands;
            _previous = commands.CommandNotFoundAction;
            _handler = OnLookup;
            commands.CommandNotFoundAction = _handler;
        }

        internal void Unregister() {
            _active = false;
            if (ReferenceEquals(_commands.CommandNotFoundAction, _handler)) _commands.CommandNotFoundAction = _previous;
        }

        private void OnLookup(object? sender, CommandLookupEventArgs eventArgs) {
            _previous?.Invoke(sender, eventArgs);
            if (!_active || eventArgs.Command != null) return;
            if (!Hints.Value.TryGetValue(eventArgs.CommandName, out string? hint)) return;
            // The lookup runs on the pipeline thread, so the warning reaches the caller's warning stream.
            _warning ??= ScriptBlock.Create("param($Message) Microsoft.PowerShell.Utility\\Write-Warning -Message $Message");
            _warning.Invoke(eventArgs.CommandName + " is no longer available. " + hint);
        }
    }
}
