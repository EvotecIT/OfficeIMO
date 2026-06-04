using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Save-time helper methods for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {

        private static void FixContentTypes(string filePath, int masterCount, bool includeTheme, bool includeComments, IEnumerable<string> pagePartNames) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be null or whitespace.", nameof(filePath));
            }

            if (pagePartNames is null) {
                throw new ArgumentNullException(nameof(pagePartNames));
            }

            using FileStream zipStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Update);
            FixContentTypesCore(archive, masterCount, includeTheme, includeComments, pagePartNames);
        }

        private static void FixContentTypes(Stream stream, int masterCount, bool includeTheme, bool includeComments, IEnumerable<string> pagePartNames) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }
            if (!stream.CanRead || !stream.CanWrite || !stream.CanSeek) {
                throw new ArgumentException("Stream must be readable, writable, and seekable.", nameof(stream));
            }
            if (pagePartNames is null) {
                throw new ArgumentNullException(nameof(pagePartNames));
            }

            stream.Seek(0, SeekOrigin.Begin);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update, leaveOpen: true);
            FixContentTypesCore(archive, masterCount, includeTheme, includeComments, pagePartNames);
            stream.Seek(0, SeekOrigin.Begin);
        }

        private static void FixContentTypesCore(ZipArchive archive, int masterCount, bool includeTheme, bool includeComments, IEnumerable<string> pagePartNames) {
            ZipArchiveEntry? entry = archive.GetEntry("[Content_Types].xml");
            entry?.Delete();
            ZipArchiveEntry newEntry = archive.CreateEntry("[Content_Types].xml");
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement root = new(ct + "Types",
                new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "emf"), new XAttribute("ContentType", "image/x-emf")),
                new XElement(ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")),
                new XElement(ct + "Default", new XAttribute("Extension", "jpg"), new XAttribute("ContentType", "image/jpeg")),
                new XElement(ct + "Default", new XAttribute("Extension", "jpeg"), new XAttribute("ContentType", "image/jpeg")),
                new XElement(ct + "Default", new XAttribute("Extension", "gif"), new XAttribute("ContentType", "image/gif")),
                new XElement(ct + "Default", new XAttribute("Extension", "svg"), new XAttribute("ContentType", "image/svg+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "tif"), new XAttribute("ContentType", "image/tiff")),
                new XElement(ct + "Default", new XAttribute("Extension", "tiff"), new XAttribute("ContentType", "image/tiff")));

            HashSet<string> overridePartNames = new(StringComparer.OrdinalIgnoreCase);
            void AddOverride(string partName, string contentType) {
                if (string.IsNullOrWhiteSpace(partName)) {
                    return;
                }

                string normalizedPartName = NormalizePartName(partName);

                if (overridePartNames.Add(normalizedPartName)) {
                    root.Add(new XElement(ct + "Override",
                        new XAttribute("PartName", normalizedPartName),
                        new XAttribute("ContentType", contentType)));
                }
            }

            AddOverride("/visio/document.xml", DocumentContentType);
            AddOverride("/visio/pages/pages.xml", PagesContentType);
            AddOverride("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
            AddOverride("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            AddOverride("/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
            AddOverride("/docProps/thumbnail.emf", "image/x-emf");
            AddOverride("/visio/windows.xml", WindowsContentType);

            foreach (string partName in pagePartNames) {
                AddOverride(partName, PageContentType);
            }
            if (includeTheme) {
                AddOverride("/visio/theme/theme1.xml", ThemeContentType);
            }
            if (includeComments) {
                AddOverride("/visio/comments.xml", CommentsContentType);
            }
            if (masterCount > 0) {
                AddOverride("/visio/masters/masters.xml", "application/vnd.ms-visio.masters+xml");
                for (int i = 1; i <= masterCount; i++) {
                    AddOverride($"/visio/masters/master{i}.xml", "application/vnd.ms-visio.master+xml");
                }
            }
            XDocument doc = new(root);
            using StreamWriter writer = new(newEntry.Open());
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static string NormalizePartName(string partName) {
            if (partName is null) {
                throw new ArgumentNullException(nameof(partName));
            }

            return "/" + partName.TrimStart('/');
        }
    }
}
