using System.Data.Common;

namespace OfficeIMO.Data {

    /// <summary>Qualifies typed getters whose results match the ordinary mapping conversion for a current field.</summary>
    internal interface IDataReaderTypedMappingCompatibility {
        bool CanUseTypedGetter(int ordinal, Type targetType);

        /// <summary>Qualifies the same conversion and proves the current field is not DBNull.</summary>
        bool CanUseNonNullTypedGetter(int ordinal, Type targetType);
    }

    /// <summary>
    /// Qualifies a fixed set of fields for each row in one call. Readers resolve the target
    /// types once; each row then gives the same answer as calling
    /// <see cref="IDataReaderTypedMappingCompatibility.CanUseNonNullTypedGetter"/> for every field.
    /// </summary>
    internal interface IDataReaderTypedRowQualification : IDataReaderTypedMappingCompatibility {
        /// <summary>Resolves reader-owned state for the fields, or returns null to keep per-field qualification.</summary>
        object? PrepareNonNullTypedRowQualification(int[] ordinals, Type[] targetTypes);

        /// <summary>Qualifies every prepared field of the current row.</summary>
        bool CanUseNonNullTypedGetters(object preparedState);
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
