namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Adds a text-bearing shape with the specified preset geometry.
        /// </summary>
        public PowerPointTextBox AddTextShape(
            OfficePresetShapeType shapeType,
            string text,
            OfficeLength? left = null,
            OfficeLength? top = null,
            OfficeLength? width = null,
            OfficeLength? height = null,
            string? name = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            ArgumentNullException.ThrowIfNull(text);

            string preset = shapeType.ToOpenXmlToken();
            return BuildTextShape(text, preset, name ?? GenerateUniqueName(shapeType.ToString()),
                nativeLeft, nativeTop, nativeWidth, nativeHeight);
        }

    }
}
