namespace OfficeIMO.Workflows;

/// <summary>Converts captured input with an adapter-owned settings snapshot. Streams remain runner-owned.</summary>
/// <typeparam name="TSettings">Settings type accepted by the registered format owner.</typeparam>
public delegate OfficeWorkflowConversionEvidence OfficeWorkflowConfiguredConverter<TSettings>(
    Stream input, Stream output, OfficeOperationLimits limits, TSettings? settings, CancellationToken cancellationToken)
    where TSettings : class, IOfficeConversionSettings;
