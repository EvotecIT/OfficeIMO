using System.Data.Common;

namespace OfficeIMO.Data {

    /// <summary>Qualifies typed getters whose results match the ordinary mapping conversion for a current field.</summary>
    internal interface IDataReaderTypedMappingCompatibility {
        bool CanUseTypedGetter(int ordinal, Type targetType);

        /// <summary>Qualifies the same conversion and proves the current field is not DBNull.</summary>
        bool CanUseNonNullTypedGetter(int ordinal, Type targetType);
    }

    internal static class DataReaderTypedMappingCompatibility {
        /// <summary>Proves the same typed conversion and a non-null current field before snapshot capture.</summary>
        internal static bool CanUseNonNullTypedGetter(DbDataReader reader, int ordinal, Type targetType) =>
            reader is IDataReaderFastMappingValues fastValues &&
            (reader is IDataReaderTypedMappingCompatibility compatibility
                ? compatibility.CanUseNonNullTypedGetter(ordinal, targetType)
                : fastValues.HasOnlyNonNullFastValues);

        internal static bool CanUseTypedGetter(DbDataReader reader, int ordinal, Type targetType) =>
            reader is IDataReaderFastMappingValues &&
            (reader is not IDataReaderTypedMappingCompatibility compatibility ||
                compatibility.CanUseTypedGetter(ordinal, targetType));
    }
}
