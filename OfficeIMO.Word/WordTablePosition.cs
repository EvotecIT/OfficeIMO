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
        public VerticalAnchorValues? VerticalAnchor {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.VerticalAnchor != null)
                    return tableProperties.TablePositionProperties.VerticalAnchor.Value;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.VerticalAnchor = value;
                else
                    tableProperties.TablePositionProperties.VerticalAnchor = null;
            }
        }

        /// <summary>
        ///     Get or set Table Horizontal Anchor
        /// </summary>
        public HorizontalAnchorValues? HorizontalAnchor {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.HorizontalAnchor != null)
                    return tableProperties.TablePositionProperties.HorizontalAnchor.Value;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.HorizontalAnchor = value;
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
        public VerticalAlignmentValues? TablePositionYAlignment {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionYAlignment != null)
                    return tableProperties.TablePositionProperties.TablePositionYAlignment;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionYAlignment = value;
                else
                    tableProperties.TablePositionProperties.TablePositionYAlignment = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Horizontal Alignment From Anchor
        /// </summary>
        public HorizontalAlignmentValues? TablePositionXAlignment {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionXAlignment != null)
                    return tableProperties.TablePositionProperties.TablePositionXAlignment;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionXAlignment = value;
                else
                    tableProperties.TablePositionProperties.TablePositionXAlignment = null;
            }
        }

        /// <summary>
        ///     Gets or sets Table Overlap
        /// </summary>
        public TableOverlapValues? TableOverlap {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TableOverlap?.Val != null)
                    return tableProperties.TableOverlap.Val.Value;

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TableOverlap == null) tableProperties.TableOverlap = new TableOverlap();
                if (value != null)
                    tableProperties.TableOverlap.Val = value;
                else
                    tableProperties.TableOverlap.Remove();
            }
        }
    }
}
