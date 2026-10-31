using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Reuses the editable model's scalar-property projection without materializing its source rows.</summary>
    internal static (string[] Headers, Action<ExcelDocument.ExcelTabularRowWriter, T> WriteRow)
        CreateStreamingRowProjection<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>() {
        var plan = GetSimpleObjectExportPlan(typeof(T));
        if (!plan.CanUseDirectSave)
            throw new ArgumentException("Rows must expose readable, non-indexed scalar properties. Supply headers and a row callback for another schema.");
        return (plan.Headers, (writer, row) => {
            if (row == null) throw new ArgumentException("A table row must not be null.", nameof(row));
            foreach (var getter in plan.Getters) writer.Write(getter(row));
        });
    }
}
