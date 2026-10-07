using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    // Value storage avoids per-pixel allocations for the bounded eight-channel input path.
    private readonly struct DeviceComponentValues : IReadOnlyList<double> {
        private readonly double _component0;
        private readonly double _component1;
        private readonly double _component2;
        private readonly double _component3;
        private readonly double _component4;
        private readonly double _component5;
        private readonly double _component6;
        private readonly double _component7;
        internal DeviceComponentValues(int count, double component0, double component1, double component2,
            double component3, double component4 = 0D, double component5 = 0D, double component6 = 0D, double component7 = 0D) {
            Count = count;
            _component0 = component0;
            _component1 = component1;
            _component2 = component2;
            _component3 = component3;
            _component4 = component4;
            _component5 = component5;
            _component6 = component6;
            _component7 = component7;
        }

        internal DeviceComponentValues(IReadOnlyList<double> components, int count)
            : this(count, components[0], components[1], components[2],
                count > 3 ? components[3] : 0D, count > 4 ? components[4] : 0D,
                count > 5 ? components[5] : 0D, count > 6 ? components[6] : 0D, count > 7 ? components[7] : 0D) { }
        public int Count { get; }
        public double this[int index] => index switch {
            0 when Count > 0 => _component0,
            1 when Count > 1 => _component1,
            2 when Count > 2 => _component2,
            3 when Count > 3 => _component3,
            4 when Count > 4 => _component4,
            5 when Count > 5 => _component5,
            6 when Count > 6 => _component6,
            7 when Count > 7 => _component7,
            _ => throw new ArgumentOutOfRangeException(nameof(index))
        };
        internal double[] ToArray() { var values = new double[Count]; CopyTo(values); return values; }
        internal void CopyTo(double[] destination) {
            for (int index = 0; index < Count; index++) destination[index] = this[index];
        }
        public IEnumerator<double> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
