using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Adds an editable native table with common header and naming options.</summary>
    public PowerPointTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null) {
        var table = AddTable(rows, configure: null, includeHeaders: options?.IncludeHeaders ?? true);
        if (options?.Name != null) table.Name = options.Name;
        return table;
    }
}
