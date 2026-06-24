namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Shared decoders for common-object metadata stored in legacy XLS OBJ records.
    /// </summary>
    internal static class LegacyXlsDrawingObjectMetadata {
        internal static LegacyXlsDrawingObjectType? TryGetObjectTypeKind(ushort? objectType) {
            if (!objectType.HasValue) {
                return null;
            }

            return objectType.Value switch {
                0x0000 => LegacyXlsDrawingObjectType.Group,
                0x0001 => LegacyXlsDrawingObjectType.Line,
                0x0002 => LegacyXlsDrawingObjectType.Rectangle,
                0x0003 => LegacyXlsDrawingObjectType.Oval,
                0x0004 => LegacyXlsDrawingObjectType.Arc,
                0x0005 => LegacyXlsDrawingObjectType.Chart,
                0x0006 => LegacyXlsDrawingObjectType.Text,
                0x0007 => LegacyXlsDrawingObjectType.Button,
                0x0008 => LegacyXlsDrawingObjectType.Picture,
                0x0009 => LegacyXlsDrawingObjectType.Polygon,
                0x000B => LegacyXlsDrawingObjectType.Checkbox,
                0x000C => LegacyXlsDrawingObjectType.RadioButton,
                0x000D => LegacyXlsDrawingObjectType.EditBox,
                0x000E => LegacyXlsDrawingObjectType.Label,
                0x000F => LegacyXlsDrawingObjectType.DialogBox,
                0x0010 => LegacyXlsDrawingObjectType.SpinControl,
                0x0011 => LegacyXlsDrawingObjectType.Scrollbar,
                0x0012 => LegacyXlsDrawingObjectType.List,
                0x0013 => LegacyXlsDrawingObjectType.GroupBox,
                0x0014 => LegacyXlsDrawingObjectType.DropdownList,
                0x0019 => LegacyXlsDrawingObjectType.Note,
                0x001E => LegacyXlsDrawingObjectType.OfficeArtObject,
                _ => null
            };
        }

        internal static IReadOnlyList<string> GetObjectFlagNames(ushort flags) {
            var names = new List<string>();
            if ((flags & 0x0001) != 0) names.Add("Locked");
            if ((flags & 0x0004) != 0) names.Add("DefaultSize");
            if ((flags & 0x0008) != 0) names.Add("Published");
            if ((flags & 0x0010) != 0) names.Add("Printable");
            if ((flags & 0x0080) != 0) names.Add("Disabled");
            if ((flags & 0x0100) != 0) names.Add("UiObject");
            if ((flags & 0x0200) != 0) names.Add("RecalculateOnLoad");
            if ((flags & 0x1000) != 0) names.Add("AlwaysRecalculate");
            return names;
        }
    }
}
