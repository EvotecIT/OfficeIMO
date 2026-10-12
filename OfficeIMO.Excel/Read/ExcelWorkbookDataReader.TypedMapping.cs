namespace OfficeIMO.Excel {
    public sealed partial class ExcelWorkbookDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
        // Worksheet shapes can contain missing cells or mixed source types. Keep qualification per row.
        bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

        bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) =>
            _current is IDataReaderTypedMappingCompatibility compatibility &&
            compatibility.CanUseTypedGetter(ordinal, targetType);
    }
}
