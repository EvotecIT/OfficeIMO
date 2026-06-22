using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel document and provides methods for creating,
    /// loading and saving spreadsheets.
    /// </summary>
    public partial class ExcelDocument {
        private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

        /// <summary>
        /// Ensures the workbook contains the default Office theme and stylesheet parts.
        /// </summary>
        public ExcelDocument EnsureWorkbookTheme() {
            EnsureWorkbookThemeAndStyles();
            return this;
        }

        /// <summary>
        /// Replaces the workbook theme with the embedded default Office theme.
        /// </summary>
        public ExcelDocument ResetWorkbookTheme(string? name = null) {
            string xml = Encoding.UTF8.GetString(DefaultThemeBytes.Value);
            if (!string.IsNullOrWhiteSpace(name)) {
                xml = RenameThemeXml(xml, name!);
            }

            SetWorkbookThemeXml(xml);
            return this;
        }

        /// <summary>
        /// Replaces or creates the workbook theme part with caller-supplied DrawingML theme XML.
        /// </summary>
        public ExcelDocument SetWorkbookThemeXml(string themeXml) {
            if (string.IsNullOrWhiteSpace(themeXml)) {
                throw new ArgumentException("Theme XML cannot be null or empty.", nameof(themeXml));
            }

            XDocument.Parse(themeXml);

            var workbookPart = _spreadSheetDocument?.WorkbookPart ?? _workBookPart;
            ThemePart themePart = workbookPart.GetPartsOfType<ThemePart>().FirstOrDefault() ?? workbookPart.AddNewPart<ThemePart>();
            byte[] bytes = Encoding.UTF8.GetBytes(themeXml);
            using var stream = new MemoryStream(bytes);
            themePart.FeedData(stream);

            EnsureWorkbookThemeAndStyles();
            Save();
            return this;
        }

        /// <summary>
        /// Sets the workbook theme name while preserving the rest of the theme XML.
        /// </summary>
        public ExcelDocument SetWorkbookThemeName(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Theme name cannot be null or empty.", nameof(name));
            }

            string xml = GetWorkbookThemeXml() ?? Encoding.UTF8.GetString(DefaultThemeBytes.Value);
            SetWorkbookThemeXml(RenameThemeXml(xml, name));
            return this;
        }

        /// <summary>
        /// Gets workbook theme metadata and optionally includes the raw theme XML.
        /// </summary>
        public ExcelWorkbookThemeInfo GetWorkbookTheme(bool includeXml = false) {
            string? xml = GetWorkbookThemeXml();
            if (xml == null) {
                return new ExcelWorkbookThemeInfo(false, null, null);
            }

            return new ExcelWorkbookThemeInfo(true, TryGetThemeName(xml), includeXml ? xml : null);
        }

        /// <summary>
        /// Gets the raw workbook theme XML when a theme part exists.
        /// </summary>
        public string? GetWorkbookThemeXml() {
            var workbookPart = _spreadSheetDocument?.WorkbookPart ?? _workBookPart;
            ThemePart? themePart = workbookPart.GetPartsOfType<ThemePart>().FirstOrDefault();
            if (themePart == null) {
                return null;
            }

            using Stream stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }

        private static string RenameThemeXml(string themeXml, string name) {
            XDocument document = XDocument.Parse(themeXml);
            XElement root = document.Root ?? throw new InvalidOperationException("Theme XML did not contain a root element.");
            if (root.Name != DrawingNamespace + "theme") {
                throw new InvalidOperationException("Theme XML root must be a DrawingML theme element.");
            }

            root.SetAttributeValue("name", name);
            return document.ToString(SaveOptions.DisableFormatting);
        }

        private static string? TryGetThemeName(string themeXml) {
            try {
                XDocument document = XDocument.Parse(themeXml);
                return document.Root?.Attribute("name")?.Value;
            } catch {
                return null;
            }
        }
    }
}
