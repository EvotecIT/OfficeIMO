namespace OfficeIMO.Workflows;

/// <summary>Typed settings owned by an opt-in conversion adapter.</summary>
public interface IOfficeWorkflowConversionSettings {
    /// <summary>Validates and creates an independent settings snapshot before asynchronous execution.</summary>
    IOfficeWorkflowConversionSettings Snapshot();
}

/// <summary>Converts captured input with an adapter-owned settings snapshot. Streams remain runner-owned.</summary>
/// <typeparam name="TSettings">Settings type accepted by the registered format owner.</typeparam>
public delegate OfficeWorkflowConversionEvidence OfficeWorkflowConfiguredConverter<TSettings>(
    Stream input, Stream output, OfficeWorkflowLimits limits, TSettings? settings, CancellationToken cancellationToken)
    where TSettings : class, IOfficeWorkflowConversionSettings;
