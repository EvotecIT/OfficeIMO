using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Adds an editable native table with common header and naming options.</summary>
    /// <remarks><see cref="OfficeTableOptions.Name"/> becomes the graphic-frame name. PowerPoint tables have no visible caption, so a non-null <see cref="OfficeTableOptions.Caption"/> throws <see cref="NotSupportedException"/>.</remarks>
    public PowerPointTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null) {
        OfficeIMO.Core.Internal.ObjectDataHelpers.ValidateTableAlignments(options);
        if (options?.Caption != null)
            throw new NotSupportedException("PowerPoint tables have no visible caption. Add a text box next to the table, or use OfficeTableOptions.Name for the graphic-frame name.");
        var table = AddTable(rows, configure: projection => OfficeIMO.Core.Internal.ObjectDataHelpers.ConfigureTableProjection(projection, options),
            includeHeaders: options?.IncludeHeaders ?? true);
        if (options?.Name != null) table.Name = options.Name;
        ApplyTableAlignments(table, options);
        return table;
    }
}
