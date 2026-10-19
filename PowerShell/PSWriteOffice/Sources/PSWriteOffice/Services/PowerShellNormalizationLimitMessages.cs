using System;
using System.Globalization;

namespace PSWriteOffice.Services;

/// <summary>Builds actionable messages for caller-configurable normalization limits.</summary>
internal static class PowerShellNormalizationLimitMessages
{
    internal static string Collection(string subject, int configuredMaximum, int observedMinimum)
    {
        return $"The {subject} contains at least {observedMinimum.ToString(CultureInfo.InvariantCulture)} items, " +
            $"which exceeds -MaxCollectionItems {configuredMaximum.ToString(CultureInfo.InvariantCulture)}. " +
            $"Rerun with -MaxCollectionItems {observedMinimum.ToString(CultureInfo.InvariantCulture)} or higher.";
    }

    internal static string Nesting(int configuredMaximum, int requiredMinimum)
    {
        return $"The cell value requires at least {requiredMinimum.ToString(CultureInfo.InvariantCulture)} normalization levels, " +
            $"which exceeds -MaxNestingDepth {configuredMaximum.ToString(CultureInfo.InvariantCulture)}. " +
            $"Rerun with -MaxNestingDepth {requiredMinimum.ToString(CultureInfo.InvariantCulture)} or higher.";
    }
}

/// <summary>Identifies a caller-configurable normalization limit independently from property getter failures.</summary>
internal sealed class PowerShellNormalizationLimitException : InvalidOperationException
{
    internal PowerShellNormalizationLimitException(string message) : base(message)
    {
    }
}
