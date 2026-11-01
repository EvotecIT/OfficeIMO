using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Retains native region choices independently of the shortened text replayed by a renderer.</summary>
internal sealed record OfficeRichTextFlowPlan(
    double AvailableWidth,
    IReadOnlyList<(int SourceEnd, double? FittingWidth)> Lines,
    IReadOnlyList<OfficeTextTabPlacement> Tabs);

/// <summary>A tab's exclusive source endpoint, measured advance and index in the paragraph's explicit stops; -1 denotes the default grid.</summary>
internal readonly record struct OfficeTextTabPlacement(int SourceEnd, double Advance, int StopIndex);
