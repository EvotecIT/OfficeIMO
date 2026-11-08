using OfficeIMO.Data;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSchemaInferenceDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
        bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

        // Sample replay retains its own objects and numeric serials. Only live
        // rows can use the inner reader's existing per-row getter qualification.
        bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) =>
            !_closed && !IsSampledRow
            && _inner is IDataReaderTypedMappingCompatibility compatibility
            && compatibility.CanUseTypedGetter(ordinal, targetType);

        bool IDataReaderTypedMappingCompatibility.CanUseNonNullTypedGetter(int ordinal, Type targetType) =>
            !_closed && !IsSampledRow
            && _inner is IDataReaderTypedMappingCompatibility compatibility
            && compatibility.CanUseNonNullTypedGetter(ordinal, targetType);
    }
}
