namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Adds a reusable typed component through the canonical flow engine, with optional layout constraints and position capture.
    /// </summary>
    public PdfDocument Component(
        IPdfComponent component,
        PdfFlowOptions? options = null,
        PdfLayoutPositionCapture? capture = null) {
        Guard.NotNull(component, nameof(component));
        return Flow(component.Compose, options, capture);
    }

    /// <summary>
    /// Adds a reusable page-aware component through the existing deferred-flow replay path.
    /// </summary>
    /// <remarks>
    /// The component does not introduce a separate measurement or layout engine. It is materialized
    /// by <see cref="Deferred"/> and may be invoked more than once while pagination stabilizes.
    /// </remarks>
    public PdfDocument Component(
        IPdfContextComponent component,
        PdfFlowOptions? options = null,
        PdfLayoutPositionCapture? capture = null) {
        Guard.NotNull(component, nameof(component));
        return Deferred(
            context => content => component.Compose(content, context),
            options,
            capture);
    }

    /// <summary>Adds a nested flow group with optional constraints and position capture.</summary>
    public PdfDocument Flow(
        Action<PdfItemCompose> compose,
        PdfFlowOptions? options = null,
        PdfLayoutPositionCapture? capture = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new FlowBlock(BuildFlowBlocks(compose), options, capture));
        return this;
    }

    /// <summary>Adds replayable content materialized from the live page context. Identical contexts are reused across layout stabilization passes.</summary>
    public PdfDocument Deferred(
        Func<PdfFlowContext, Action<PdfItemCompose>> composeFactory,
        PdfFlowOptions? options = null,
        PdfLayoutPositionCapture? capture = null) {
        Guard.NotNull(composeFactory, nameof(composeFactory));
        var generatedSectionOwner = new object();
        AddBlock(new FlowBlock(
            context => {
                using GeneratedSectionMaterializationScope generatedSections = BeginGeneratedSectionMaterialization(generatedSectionOwner);
                Action<PdfItemCompose> compose = composeFactory(context)
                    ?? throw new InvalidOperationException("Deferred PDF flow factory returned null.");
                return BuildFlowBlocks(compose);
            },
            options,
            capture));
        return this;
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock> BuildFlowBlocks(Action<PdfItemCompose> compose) {
        var blocks = new List<IPdfBlock>();
        using (PushBlockScope(blocks.Add)) {
            compose(new PdfItemCompose(this));
        }

        return blocks.AsReadOnly();
    }
}
