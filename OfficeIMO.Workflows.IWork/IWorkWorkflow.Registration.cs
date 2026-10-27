namespace OfficeIMO.Workflows.IWork;

public static partial class IWorkWorkflow {
    private static readonly Lazy<bool> ConverterRegistration = new(() => {
        foreach (OfficeWorkflowConversionRegistration registration in CreateRegistrations())
            OfficeConverter.Register(OfficeWorkflow.CreateConversionRoute(registration));
        return true;
    });

    /// <summary>Registers bounded Pages, Numbers and Keynote conversions with <see cref="OfficeConverter"/>. Per-operation options use <see cref="IWorkWorkflowSettings"/>.</summary>
    public static void RegisterConversions() => _ = ConverterRegistration.Value;
}
