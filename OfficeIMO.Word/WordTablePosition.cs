using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    ///     Properties for Table Positioning
    /// </summary>
    public class WordTablePosition {
        private readonly WordTable _table;

        private TableProperties? TableProperties => _table._tableProperties;

        private TableProperties EnsureTableProperties() {
            _table.CheckTableProperties();
            return _table._tableProperties ?? throw new InvalidOperationException("Table properties are not available.");
        }

        /// <summary>
        ///     Constructor for Table Positioning
        /// </summary>
        /// <param name="table"></param>
        internal WordTablePosition(WordTable table) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
        }

        /// <summary>
        ///     Get or set Distance From Left of Table to Text
        /// </summary>
        public short? LeftFromText {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.LeftFromText != null)
                    return tableProperties.TablePositionProperties.LeftFromText;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) tableProperties.TablePositionProperties.LeftFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Right of Table to Text
        /// </summary>
        public short? RightFromText {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.RightFromText != null)
                    return tableProperties.TablePositionProperties.RightFromText;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) tableProperties.TablePositionProperties.RightFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Bottom of Table to Text
        /// </summary>
        public short? BottomFromText {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.BottomFromText != null)
                    return tableProperties.TablePositionProperties.BottomFromText;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) tableProperties.TablePositionProperties.BottomFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Top of Table to Text
        /// </summary>
        public short? TopFromText {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TopFromText != null)
                    return tableProperties.TablePositionProperties.TopFromText;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) tableProperties.TablePositionProperties.TopFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Table Vertical Anchor
        /// </summary>
        public WordTableVerticalAnchor? VerticalAnchor {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.VerticalAnchor != null)
                    return tableProperties.TablePositionProperties.VerticalAnchor.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.VerticalAnchor = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.VerticalAnchor = null;
            }
        }

        /// <summary>
        ///     Get or set Table Horizontal Anchor
        /// </summary>
        public WordTableHorizontalAnchor? HorizontalAnchor {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.HorizontalAnchor != null)
                    return tableProperties.TablePositionProperties.HorizontalAnchor.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.HorizontalAnchor = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.HorizontalAnchor = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Vertical Alignment from Anchor
        /// </summary>
        public int? TablePositionY {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionY != null)
                    return tableProperties.TablePositionProperties.TablePositionY;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionY = value;
                else
                    tableProperties.TablePositionProperties.TablePositionY = null;
            }
        }

        /// <summary>
        ///     Get or set Absolute Horizontal Distance From Anchor
        /// </summary>
        public int? TablePositionX {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionX != null)
                    return tableProperties.TablePositionProperties.TablePositionX;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionX = value;
                else
                    tableProperties.TablePositionProperties.TablePositionX = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Vertical Alignment from Anchor
        /// </summary>
        public WordTableVerticalPositionAlignment? TablePositionYAlignment {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionYAlignment != null)
                    return tableProperties.TablePositionProperties.TablePositionYAlignment.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionYAlignment = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.TablePositionYAlignment = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Horizontal Alignment From Anchor
        /// </summary>
        public WordTableHorizontalAlignment? TablePositionXAlignment {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionXAlignment != null)
                    return tableProperties.TablePositionProperties.TablePositionXAlignment.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionXAlignment = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.TablePositionXAlignment = null;
            }
        }

        /// <summary>
        ///     Gets or sets Table Overlap
        /// </summary>
        public WordTableOverlap? TableOverlap {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TableOverlap?.Val != null)
                    return tableProperties.TableOverlap.Val.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TableOverlap == null) tableProperties.TableOverlap = new TableOverlap();
                if (value != null)
                    tableProperties.TableOverlap.Val = value.Value.ToOpenXml();
                else
                    tableProperties.TableOverlap.Remove();
            }
        }
    }
}
