namespace OfficeIMO.Workflows;

/// <summary>A user-selected provider destination with verified direct writes and durable recovery.</summary>
/// <remarks>Only Replace is supported. A provider cannot guarantee atomic replacement, exclusive creation, or rollback.</remarks>
public sealed class OfficeWorkflowStreamOutput {
    /// <summary>Creates a destination. The host must explain and obtain consent for direct provider writes.</summary>
    public OfficeWorkflowStreamOutput(string name, Func<CancellationToken, Task<Stream>> openRead,
        Func<CancellationToken, Task<Stream>> openWrite, OfficeWorkflowOutputRecoveryStore recoveryStore,
        Func<CancellationToken, Task<string>>? prepareDestination = null) {
        _ = new OfficeWorkflowStreamInput(name, openRead);
        OfficeWorkflowOutputRecoveryStore.ValidateExtension(name);
        Name = name;
        OpenRead = openRead;
        OpenWrite = openWrite ?? throw new ArgumentNullException(nameof(openWrite));
        RecoveryStore = recoveryStore ?? throw new ArgumentNullException(nameof(recoveryStore));
        PrepareDestination = prepareDestination;
    }
    /// <summary>Gets the selected filename, including its output format extension.</summary>
    public string Name { get; }
    /// <summary>Gets the factory used to establish local provider access and verify committed contents after closing the write stream.</summary>
    /// <remarks>For a local provider path, this may be called before writing so the host can inspect the destination while access is active.
    /// Throw <see cref="FileNotFoundException"/> when the selected new file does not exist yet. Other access failures prevent publication.</remarks>
    public Func<CancellationToken, Task<Stream>> OpenRead { get; }
    /// <summary>Gets the factory that opens a destructive write stream. The runner closes every returned stream.</summary>
    public Func<CancellationToken, Task<Stream>> OpenWrite { get; }
    /// <summary>Gets the durable store used before any destination write begins.</summary>
    public OfficeWorkflowOutputRecoveryStore RecoveryStore { get; }

    /// <summary>Gets optional deferred creation for a new provider child, returning its actual location.</summary>
    /// <remarks>Runs after durable recovery and initial authorization. Creation may modify the provider;
    /// any subsequent failure retains recovery. The returned location is authorized before opening the write stream.
    /// The host must resolve existing children without mutation and use this callback only for creation of the selected new child.</remarks>
    public Func<CancellationToken, Task<string>>? PrepareDestination { get; }
}
