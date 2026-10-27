using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Adds an editable native table with common header and naming options.</summary>
    public PowerPointTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null) {
        OfficeIMO.Core.Internal.ObjectDataHelpers.ValidateTableAlignments(options);
        var table = AddTable(rows, configure: projection => OfficeIMO.Core.Internal.ObjectDataHelpers.ConfigureTableProjection(projection, options),
            includeHeaders: options?.IncludeHeaders ?? true);
        if (options?.Name != null) table.Name = options.Name;
        ApplyTableAlignments(table, options);
        return table;
    }
}
