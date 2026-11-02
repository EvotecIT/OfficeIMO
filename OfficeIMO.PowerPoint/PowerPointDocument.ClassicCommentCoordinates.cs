using System;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        // Classic p:pos uses eighth-points, matching binary PPT's 576 master units/inch.
        // Modern comment coordinates use EMUs and do not pass through this boundary.
        /// <summary>Converts native classic coordinates, rounding half EMUs away from zero.</summary>
        internal static OfficeLength ClassicCommentPositionToLength(long value) =>
            OfficeLength.Emus(checked((long)decimal.Round(value * 12700m / 8m,
                0, MidpointRounding.AwayFromZero)));

        /// <summary>Quantizes a physical length to classic eighth-points before native range validation.</summary>
        internal static long LengthToClassicCommentPosition(OfficeLength value) =>
            checked((long)decimal.Round(value.ToEmus() * 8m / 12700m,
                0, MidpointRounding.AwayFromZero));
    }
}
