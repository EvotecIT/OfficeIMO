using System.Data.Common;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelWorkbookDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility, IDataReaderMappingRowSource {
        // Typed mapping reads the open worksheet reader directly after each successful Read.
        DbDataReader? IDataReaderMappingRowSource.CurrentRowReader => _closed ? null : _activeReader;

        // Worksheet shapes can contain missing cells or mixed source types. Keep qualification per row.
        bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

        bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) =>
            CurrentReader is IDataReaderTypedMappingCompatibility compatibility &&
            compatibility.CanUseTypedGetter(ordinal, targetType);

        bool IDataReaderTypedMappingCompatibility.CanUseNonNullTypedGetter(int ordinal, Type targetType) =>
            CurrentReader is IDataReaderTypedMappingCompatibility compatibility &&
            compatibility.CanUseNonNullTypedGetter(ordinal, targetType);
    }
}
