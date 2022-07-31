using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    ///     Properties for Table Positioning
    /// </summary>
    public class WordTablePosition {
        private readonly WordTable _table;
        private readonly TableProperties _tableProperties;

        /// <summary>
        ///     Constructor for Table Positioning
        /// </summary>
        /// <param name="table"></param>
        internal WordTablePosition(WordTable table) {
            _table = table;
            _tableProperties = table._tableProperties;
        }

        /// <summary>
        ///     Get or set Distance From Left of Table to Text
        /// </summary>
        public short? LeftFromText {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.LeftFromText != null)
                        return _tableProperties.TablePositionProperties.LeftFromText;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) _tableProperties.TablePositionProperties.LeftFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Right of Table to Text
        /// </summary>
        public short? RightFromText {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.RightFromText != null)
                        return _tableProperties.TablePositionProperties.RightFromText;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) _tableProperties.TablePositionProperties.RightFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Bottom of Table to Text
        /// </summary>
        public short? BottomFromText {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.BottomFromText != null)
                        return _tableProperties.TablePositionProperties.BottomFromText;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) _tableProperties.TablePositionProperties.BottomFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Top of Table to Text
        /// </summary>
        public short? TopFromText {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.TopFromText != null)
                        return _tableProperties.TablePositionProperties.TopFromText;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null) _tableProperties.TablePositionProperties.TopFromText = value;
            }
        }

        /// <summary>
        ///     Get or set Table Vertical Anchor
        /// </summary>
        public VerticalAnchorValues? VerticalAnchor {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.VerticalAnchor != null)
                        return _tableProperties.TablePositionProperties.VerticalAnchor.Value;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    _tableProperties.TablePositionProperties.VerticalAnchor = value;
                else
                    _tableProperties.TablePositionProperties.VerticalAnchor = null;
            }
        }

        /// <summary>
        ///     Get or set Table Horizontal Anchor
        /// </summary>
        public HorizontalAnchorValues? HorizontalAnchor {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.HorizontalAnchor != null)
                        return _tableProperties.TablePositionProperties.HorizontalAnchor.Value;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    _tableProperties.TablePositionProperties.HorizontalAnchor = value;
                else
                    _tableProperties.TablePositionProperties.HorizontalAnchor = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Vertical Alignment from Anchor
        /// </summary>
        public int? TablePositionY {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.TablePositionY != null)
                        return _tableProperties.TablePositionProperties.TablePositionY;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    _tableProperties.TablePositionProperties.TablePositionY = value;
                else
                    _tableProperties.TablePositionProperties.TablePositionY = null;
            }
        }

        /// <summary>
        ///     Get or set Absolute Horizontal Distance From Anchor
        /// </summary>
        public int? TablePositionX {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.TablePositionX != null)
                        return _tableProperties.TablePositionProperties.TablePositionX;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    _tableProperties.TablePositionProperties.TablePositionX = value;
                else
                    _tableProperties.TablePositionProperties.TablePositionX = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Vertical Alignment from Anchor
        /// </summary>
        public VerticalAlignmentValues? TablePositionYAlignment {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.TablePositionYAlignment != null)
                        return _tableProperties.TablePositionProperties.TablePositionYAlignment;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    _tableProperties.TablePositionProperties.TablePositionYAlignment = value;
                else
                    _tableProperties.TablePositionProperties.TablePositionYAlignment = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Horizontal Alignment From Anchor
        /// </summary>
        public HorizontalAlignmentValues? TablePositionXAlignment {
            get {
                if (_tableProperties != null && _tableProperties.TablePositionProperties != null)
                    if (_tableProperties.TablePositionProperties.TablePositionXAlignment != null)
                        return _tableProperties.TablePositionProperties.TablePositionXAlignment;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TablePositionProperties == null)
                    _tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    _tableProperties.TablePositionProperties.TablePositionXAlignment = value;
                else
                    _tableProperties.TablePositionProperties.TablePositionXAlignment = null;
            }
        }

        /// <summary>
        ///     Gets or sets Table Overlap
        /// </summary>
        public TableOverlapValues? TableOverlap {
            get {
                if (_tableProperties != null && _tableProperties.TableOverlap != null)
                    return _tableProperties.TableOverlap.Val;

                return null;
            }
            set {
                _table.CheckTableProperties();
                if (_tableProperties.TableOverlap == null) _tableProperties.TableOverlap = new TableOverlap();
                if (value != null)
                    _tableProperties.TableOverlap.Val = value;
                else
                    _tableProperties.TableOverlap.Remove();
            }
        }
    }
}