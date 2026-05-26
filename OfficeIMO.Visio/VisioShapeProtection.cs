using System;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// ShapeSheet protection cells that control how a shape or connector can be edited in Visio.
    /// </summary>
    public class VisioProtection {
        internal static readonly string[] CellNames = {
            "LockWidth",
            "LockHeight",
            "LockAspect",
            "LockMoveX",
            "LockMoveY",
            "LockDelete",
            "LockTextEdit",
            "LockFormat",
            "LockGroup",
            "LockUngroup",
            "LockSelect",
            "LockRotate",
            "LockCrop",
            "LockVtxEdit",
            "LockBegin",
            "LockEnd",
            "LockCalcWH",
            "LockCustProp",
            "LockFromGroupFormat",
            "LockThemeColors",
            "LockThemeEffects"
        };

        /// <summary>Locks the shape width.</summary>
        public bool? LockWidth { get; set; }

        /// <summary>Locks the shape height.</summary>
        public bool? LockHeight { get; set; }

        /// <summary>Locks the shape aspect ratio.</summary>
        public bool? LockAspect { get; set; }

        /// <summary>Locks horizontal movement.</summary>
        public bool? LockMoveX { get; set; }

        /// <summary>Locks vertical movement.</summary>
        public bool? LockMoveY { get; set; }

        /// <summary>Prevents deleting the shape.</summary>
        public bool? LockDelete { get; set; }

        /// <summary>Prevents editing the shape text.</summary>
        public bool? LockTextEdit { get; set; }

        /// <summary>Prevents formatting changes.</summary>
        public bool? LockFormat { get; set; }

        /// <summary>Prevents grouping the shape.</summary>
        public bool? LockGroup { get; set; }

        /// <summary>Prevents ungrouping the shape.</summary>
        public bool? LockUngroup { get; set; }

        /// <summary>Prevents selecting the shape.</summary>
        public bool? LockSelect { get; set; }

        /// <summary>Locks shape rotation.</summary>
        public bool? LockRotate { get; set; }

        /// <summary>Locks cropping operations.</summary>
        public bool? LockCrop { get; set; }

        /// <summary>Locks vertex editing.</summary>
        public bool? LockVtxEdit { get; set; }

        /// <summary>Locks the begin endpoint for one-dimensional shapes.</summary>
        public bool? LockBegin { get; set; }

        /// <summary>Locks the end endpoint for one-dimensional shapes.</summary>
        public bool? LockEnd { get; set; }

        /// <summary>Locks recalculation of width and height.</summary>
        public bool? LockCalcWH { get; set; }

        /// <summary>Locks custom properties / Shape Data editing.</summary>
        public bool? LockCustProp { get; set; }

        /// <summary>Locks formatting inherited from a group.</summary>
        public bool? LockFromGroupFormat { get; set; }

        /// <summary>Locks theme color changes.</summary>
        public bool? LockThemeColors { get; set; }

        /// <summary>Locks theme effect changes.</summary>
        public bool? LockThemeEffects { get; set; }

        /// <summary>
        /// Gets whether any protection cell has been explicitly set.
        /// </summary>
        public bool HasAnyLocks => CellNames.Any(name => GetCellValue(name).HasValue);

        /// <summary>
        /// Locks or unlocks shape size.
        /// </summary>
        public VisioProtection Size(bool locked = true) {
            LockWidth = locked;
            LockHeight = locked;
            return this;
        }

        /// <summary>
        /// Locks or unlocks shape position.
        /// </summary>
        public VisioProtection Position(bool locked = true) {
            LockMoveX = locked;
            LockMoveY = locked;
            return this;
        }

        /// <summary>
        /// Locks or unlocks text editing.
        /// </summary>
        public VisioProtection Text(bool locked = true) {
            LockTextEdit = locked;
            return this;
        }

        /// <summary>
        /// Locks or unlocks deletion.
        /// </summary>
        public VisioProtection Deletion(bool locked = true) {
            LockDelete = locked;
            return this;
        }

        /// <summary>
        /// Locks or unlocks formatting.
        /// </summary>
        public VisioProtection Formatting(bool locked = true) {
            LockFormat = locked;
            return this;
        }

        /// <summary>
        /// Locks or unlocks selection.
        /// </summary>
        public VisioProtection Selection(bool locked = true) {
            LockSelect = locked;
            return this;
        }

        /// <summary>
        /// Locks or unlocks connector endpoints.
        /// </summary>
        public VisioProtection Endpoints(bool locked = true) {
            LockBegin = locked;
            LockEnd = locked;
            return this;
        }

        /// <summary>
        /// Clears every explicit protection setting.
        /// </summary>
        public VisioProtection Clear() {
            foreach (string cellName in CellNames) {
                SetCellValue(cellName, null);
            }

            return this;
        }

        internal static bool IsCellName(string? cellName) {
            return !string.IsNullOrWhiteSpace(cellName) &&
                   CellNames.Any(name => string.Equals(name, cellName, StringComparison.OrdinalIgnoreCase));
        }

        internal bool TryGetCellValue(string cellName, out bool? value) {
            string? canonicalName = GetCanonicalCellName(cellName);
            if (canonicalName == null) {
                value = null;
                return false;
            }

            value = GetCellValue(canonicalName);
            return true;
        }

        internal bool TrySetCellValue(string? cellName, bool? value) {
            string? canonicalName = GetCanonicalCellName(cellName);
            if (canonicalName == null) {
                return false;
            }

            SetCellValue(canonicalName, value);
            return true;
        }

        private static string? GetCanonicalCellName(string? cellName) {
            return string.IsNullOrWhiteSpace(cellName)
                ? null
                : CellNames.FirstOrDefault(name => string.Equals(name, cellName, StringComparison.OrdinalIgnoreCase));
        }

        private bool? GetCellValue(string cellName) {
            switch (cellName) {
                case "LockWidth": return LockWidth;
                case "LockHeight": return LockHeight;
                case "LockAspect": return LockAspect;
                case "LockMoveX": return LockMoveX;
                case "LockMoveY": return LockMoveY;
                case "LockDelete": return LockDelete;
                case "LockTextEdit": return LockTextEdit;
                case "LockFormat": return LockFormat;
                case "LockGroup": return LockGroup;
                case "LockUngroup": return LockUngroup;
                case "LockSelect": return LockSelect;
                case "LockRotate": return LockRotate;
                case "LockCrop": return LockCrop;
                case "LockVtxEdit": return LockVtxEdit;
                case "LockBegin": return LockBegin;
                case "LockEnd": return LockEnd;
                case "LockCalcWH": return LockCalcWH;
                case "LockCustProp": return LockCustProp;
                case "LockFromGroupFormat": return LockFromGroupFormat;
                case "LockThemeColors": return LockThemeColors;
                case "LockThemeEffects": return LockThemeEffects;
                default: return null;
            }
        }

        private void SetCellValue(string cellName, bool? value) {
            switch (cellName) {
                case "LockWidth":
                    LockWidth = value;
                    break;
                case "LockHeight":
                    LockHeight = value;
                    break;
                case "LockAspect":
                    LockAspect = value;
                    break;
                case "LockMoveX":
                    LockMoveX = value;
                    break;
                case "LockMoveY":
                    LockMoveY = value;
                    break;
                case "LockDelete":
                    LockDelete = value;
                    break;
                case "LockTextEdit":
                    LockTextEdit = value;
                    break;
                case "LockFormat":
                    LockFormat = value;
                    break;
                case "LockGroup":
                    LockGroup = value;
                    break;
                case "LockUngroup":
                    LockUngroup = value;
                    break;
                case "LockSelect":
                    LockSelect = value;
                    break;
                case "LockRotate":
                    LockRotate = value;
                    break;
                case "LockCrop":
                    LockCrop = value;
                    break;
                case "LockVtxEdit":
                    LockVtxEdit = value;
                    break;
                case "LockBegin":
                    LockBegin = value;
                    break;
                case "LockEnd":
                    LockEnd = value;
                    break;
                case "LockCalcWH":
                    LockCalcWH = value;
                    break;
                case "LockCustProp":
                    LockCustProp = value;
                    break;
                case "LockFromGroupFormat":
                    LockFromGroupFormat = value;
                    break;
                case "LockThemeColors":
                    LockThemeColors = value;
                    break;
                case "LockThemeEffects":
                    LockThemeEffects = value;
                    break;
            }
        }
    }

    /// <summary>
    /// Shape-specific view over native Visio ShapeSheet protection cells.
    /// </summary>
    public sealed class VisioShapeProtection : VisioProtection {
    }
}
