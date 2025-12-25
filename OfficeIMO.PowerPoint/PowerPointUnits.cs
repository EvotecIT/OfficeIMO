using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides helpers for converting between EMUs and common units.
    /// </summary>
    public static class PowerPointUnits {
        /// <summary>
        ///     EMUs per inch.
        /// </summary>
        public const double EmusPerInch = 914400d;

        /// <summary>
        ///     EMUs per point.
        /// </summary>
        public const double EmusPerPoint = 12700d;

        /// <summary>
        ///     EMUs per centimeter.
        /// </summary>
        public const double EmusPerCentimeter = EmusPerInch / 2.54d;

        /// <summary>
        ///     EMUs per millimeter.
        /// </summary>
        public const double EmusPerMillimeter = EmusPerCentimeter / 10d;

        /// <summary>
        ///     Converts inches to EMUs.
        /// </summary>
        public static long FromInches(double inches) => (long)Math.Round(inches * EmusPerInch);

        /// <summary>
        ///     Converts points to EMUs.
        /// </summary>
        public static long FromPoints(double points) => (long)Math.Round(points * EmusPerPoint);

        /// <summary>
        ///     Converts centimeters to EMUs.
        /// </summary>
        public static long FromCentimeters(double centimeters) => (long)Math.Round(centimeters * EmusPerCentimeter);

        /// <summary>
        ///     Converts millimeters to EMUs.
        /// </summary>
        public static long FromMillimeters(double millimeters) => (long)Math.Round(millimeters * EmusPerMillimeter);

        /// <summary>
        ///     Converts EMUs to inches.
        /// </summary>
        public static double ToInches(long emus) => emus / EmusPerInch;

        /// <summary>
        ///     Converts EMUs to points.
        /// </summary>
        public static double ToPoints(long emus) => emus / EmusPerPoint;

        /// <summary>
        ///     Converts EMUs to centimeters.
        /// </summary>
        public static double ToCentimeters(long emus) => emus / EmusPerCentimeter;

        /// <summary>
        ///     Converts EMUs to millimeters.
        /// </summary>
        public static double ToMillimeters(long emus) => emus / EmusPerMillimeter;

        // Short aliases for call sites (e.g., PowerPointUnits.Cm(2.5))

        /// <summary>
        ///     Alias for <see cref="FromInches(double)"/>.
        /// </summary>
        public static long Inches(double inches) => FromInches(inches);

        /// <summary>
        ///     Alias for <see cref="FromPoints(double)"/>.
        /// </summary>
        public static long Points(double points) => FromPoints(points);

        /// <summary>
        ///     Alias for <see cref="FromCentimeters(double)"/>.
        /// </summary>
        public static long Cm(double centimeters) => FromCentimeters(centimeters);

        /// <summary>
        ///     Alias for <see cref="FromMillimeters(double)"/>.
        /// </summary>
        public static long Mm(double millimeters) => FromMillimeters(millimeters);
    }
}
