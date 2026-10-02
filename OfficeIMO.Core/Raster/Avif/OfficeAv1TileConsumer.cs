namespace OfficeIMO.Drawing;

/// <summary>Streaming sink for tile syntax in reconstruction order.</summary>
/// <remarks>Callbacks run synchronously. A reconstruction owner budgets any retained results and
/// discards the whole frame on failure; only CompleteTile signifies successful entropy termination.
/// No callback may resume or reenter the same tile reader.</remarks>
internal interface IOfficeAv1TileConsumer {
    void Restoration(OfficeAv1RestorationUnit unit);
    void BeginBlock(OfficeAv1TileBlock block);
    void Residual(OfficeAv1Coefficients coefficients);
    void EndBlock();
    void CompleteTile();
}

/// <summary>Immutable leaf inputs retained only while its residuals are streamed to the consumer.</summary>
internal readonly struct OfficeAv1TileBlock {
    internal OfficeAv1TileBlock(OfficeAv1BlockRegion region, OfficeAv1BlockPrelude prelude,
        OfficeAv1IntraModes modes, OfficeAv1CopyMotion motion, OfficeAv1Palette palette, OfficeAv1TransformLayout transforms) {
        Region=region; Prelude=prelude; Modes=modes; Motion=motion; Palette=palette; Transforms=transforms;
    }
    internal OfficeAv1BlockRegion Region { get; }
    internal OfficeAv1BlockPrelude Prelude { get; }
    internal OfficeAv1IntraModes Modes { get; }
    internal OfficeAv1CopyMotion Motion { get; }
    internal OfficeAv1Palette Palette { get; }
    internal OfficeAv1TransformLayout Transforms { get; }
}
