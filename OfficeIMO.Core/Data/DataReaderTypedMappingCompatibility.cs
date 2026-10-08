#nullable enable

using System;
using System.Data.Common;

namespace OfficeIMO.Data;

/// <summary>Qualifies typed getters whose results match the ordinary mapping conversion for a current field.</summary>
internal interface IDataReaderTypedMappingCompatibility {
    bool CanUseTypedGetter(int ordinal, Type targetType);
}

internal static class DataReaderTypedMappingCompatibility {
    internal static bool CanUseTypedGetter(DbDataReader reader, int ordinal, Type targetType) =>
        reader is IDataReaderFastMappingValues &&
        (reader is not IDataReaderTypedMappingCompatibility compatibility ||
            compatibility.CanUseTypedGetter(ordinal, targetType));
}
