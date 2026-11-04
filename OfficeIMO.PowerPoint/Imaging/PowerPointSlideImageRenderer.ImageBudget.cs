using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
        internal sealed class PowerPointEmbeddedImageReadBudget {
            private readonly int _maximumImageBytes;
            private long _remainingBytes;

            internal PowerPointEmbeddedImageReadBudget(int maximumImageBytes, long maximumTotalBytes) {
                _maximumImageBytes = maximumImageBytes;
                _remainingBytes = maximumTotalBytes;
            }

            internal bool IsExhausted => _remainingBytes <= 0;
            internal int GetBackgroundReadLimit() => IsExhausted ? 1 : GetNextReadLimit();

            internal int GetNextReadLimit() {
                if (_remainingBytes <= 0) {
                    throw new InvalidOperationException("The aggregate embedded-image byte limit was exceeded.");
                }
                return (int)Math.Min(_maximumImageBytes, Math.Min(int.MaxValue, _remainingBytes));
            }

            internal void Consume(int bytes) {
                if (bytes < 0 || bytes > _remainingBytes) {
                    throw new InvalidOperationException("The aggregate embedded-image byte limit was exceeded.");
                }
                _remainingBytes -= bytes;
            }
        }
}
