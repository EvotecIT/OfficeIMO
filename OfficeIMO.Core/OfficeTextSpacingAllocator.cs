using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>A caller-unit spacing resource. Lower numeric priorities are consumed first.</summary>
internal readonly struct OfficeTextSpacingResource {
    internal OfficeTextSpacingResource(int opportunityIndex, int priority, double compressionCapacity, double expansionCapacity) {
        if (opportunityIndex < 0) throw new ArgumentOutOfRangeException(nameof(opportunityIndex));
        ValidateCapacity(compressionCapacity, nameof(compressionCapacity));
        ValidateCapacity(expansionCapacity, nameof(expansionCapacity));
        OpportunityIndex = opportunityIndex;
        Priority = priority;
        CompressionCapacity = compressionCapacity;
        ExpansionCapacity = expansionCapacity;
    }

    internal int OpportunityIndex { get; }
    internal int Priority { get; }
    internal double CompressionCapacity { get; }
    internal double ExpansionCapacity { get; }

    private static void ValidateCapacity(double value, string name) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
    }
}

/// <summary>An immutable allocation and its actual floating-point residual, in the caller's units.</summary>
internal sealed class OfficeTextSpacingAllocation {
    internal OfficeTextSpacingAllocation(double[] deltas, double request) {
        var snapshot = (double[])deltas.Clone();
        Deltas = Array.AsReadOnly(snapshot);
        double applied = 0D;
        foreach (double delta in snapshot) {
            applied += delta;
            if (double.IsInfinity(applied)) throw new OverflowException("The allocated spacing total is not finite.");
        }
        AppliedAdjustment = applied;
        UnfulfilledAdjustment = request - applied;
    }

    internal IReadOnlyList<double> Deltas { get; }
    internal double AppliedAdjustment { get; }
    internal double UnfulfilledAdjustment { get; }
}

/// <summary>Allocates signed spacing over explicit resources without font, shaping, or native-profile policy.</summary>
internal static class OfficeTextSpacingAllocator {
    /// <summary>
    /// Consumes lower priorities completely before later priorities. Within a partial priority,
    /// each capacity supplies its proportional share in the caller's original resource order.
    /// Floating-point remainder is carried forward; no integer quantum or fit tolerance is imposed.
    /// </summary>
    internal static OfficeTextSpacingAllocation Allocate(int opportunityCount, IReadOnlyList<OfficeTextSpacingResource> resources, double adjustment) {
        if (opportunityCount < 0) throw new ArgumentOutOfRangeException(nameof(opportunityCount));
        if (resources == null) throw new ArgumentNullException(nameof(resources));
        if (double.IsNaN(adjustment) || double.IsInfinity(adjustment)) throw new ArgumentOutOfRangeException(nameof(adjustment));

        var priorities = new SortedDictionary<int, ResourcePriority>();
        double totalCompression = 0D, totalExpansion = 0D;
        for (int index = 0; index < resources.Count; index++) {
            OfficeTextSpacingResource resource = resources[index];
            if (resource.OpportunityIndex >= opportunityCount) throw new ArgumentOutOfRangeException(nameof(resources), "A spacing opportunity is outside the result.");
            totalCompression = AddCapacity(totalCompression, resource.CompressionCapacity);
            totalExpansion = AddCapacity(totalExpansion, resource.ExpansionCapacity);
            if (!priorities.TryGetValue(resource.Priority, out ResourcePriority? priority)) {
                priority = new ResourcePriority();
                priorities.Add(resource.Priority, priority);
            }
            priority.Resources.Add(resource);
            priority.CompressionCapacity = AddCapacity(priority.CompressionCapacity, resource.CompressionCapacity);
            priority.ExpansionCapacity = AddCapacity(priority.ExpansionCapacity, resource.ExpansionCapacity);
        }

        // Validate both directional totals before creating or applying any allocation, even for zero requests.
        var deltas = new double[opportunityCount];
        bool compress = adjustment < 0D;
        double available = compress ? totalCompression : totalExpansion;
        if (adjustment == 0D || available == 0D) return new OfficeTextSpacingAllocation(deltas, adjustment);
        double remaining = Math.Abs(adjustment);
        double sign = compress ? -1D : 1D;
        foreach (ResourcePriority priority in priorities.Values) {
            double capacity = compress ? priority.CompressionCapacity : priority.ExpansionCapacity;
            if (capacity == 0D) continue;
            if (remaining >= capacity) {
                foreach (OfficeTextSpacingResource resource in priority.Resources)
                    Apply(resource.OpportunityIndex, compress ? resource.CompressionCapacity : resource.ExpansionCapacity);
                remaining -= capacity;
                if (remaining == 0D) break;
                continue;
            }

            double remainingCapacity = capacity;
            foreach (OfficeTextSpacingResource resource in priority.Resources) {
                double resourceCapacity = compress ? resource.CompressionCapacity : resource.ExpansionCapacity;
                if (resourceCapacity == 0D) continue;
                // Divide before multiplying so two finite large values need not overflow.
                double amount = remainingCapacity <= resourceCapacity
                    ? remaining : remaining * (resourceCapacity / remainingCapacity);
                amount = Math.Min(resourceCapacity, Math.Min(remaining, amount));
                Apply(resource.OpportunityIndex, amount);
                remaining -= amount;
                remainingCapacity = Math.Max(0D, remainingCapacity - resourceCapacity);
            }
            // A partial priority retains its representational residual; later priorities must not mask it.
            break;
        }
        return new OfficeTextSpacingAllocation(deltas, adjustment);

        void Apply(int index, double amount) {
            double value = deltas[index] + sign * amount;
            if (double.IsInfinity(value)) throw new OverflowException("An allocated spacing opportunity is not finite.");
            deltas[index] = value;
        }
    }

    private static double AddCapacity(double total, double capacity) {
        double sum = total + capacity;
        if (double.IsInfinity(sum)) throw new OverflowException("The spacing resource capacity total is not finite.");
        return sum;
    }

    private sealed class ResourcePriority {
        internal readonly List<OfficeTextSpacingResource> Resources = new();
        internal double CompressionCapacity;
        internal double ExpansionCapacity;
    }
}
