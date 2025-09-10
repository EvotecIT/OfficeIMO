using System;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Fluent builder for workbook properties (core + extended).
    /// </summary>
    public sealed class InfoBuilder
    {
        private readonly ExcelDocument _doc;
        internal InfoBuilder(ExcelDocument doc) { _doc = doc; }

        /// <summary>Sets core property Title.</summary>
        public InfoBuilder Title(string title) { _doc.BuiltinDocumentProperties.Title = title; return this; }
        /// <summary>Sets core property Creator/Author.</summary>
        public InfoBuilder Author(string author) { _doc.BuiltinDocumentProperties.Creator = author; return this; }
        /// <summary>Sets core property Subject.</summary>
        public InfoBuilder Subject(string subject) { _doc.BuiltinDocumentProperties.Subject = subject; return this; }
        /// <summary>Sets core property Keywords.</summary>
        public InfoBuilder Keywords(string keywords) { _doc.BuiltinDocumentProperties.Keywords = keywords; return this; }
        /// <summary>Sets core property Description.</summary>
        public InfoBuilder Description(string description) { _doc.BuiltinDocumentProperties.Description = description; return this; }
        /// <summary>Sets core property Category.</summary>
        public InfoBuilder Category(string category) { _doc.BuiltinDocumentProperties.Category = category; return this; }
        /// <summary>Sets extended property Company.</summary>
        public InfoBuilder Company(string company) { _doc.ApplicationProperties.Company = company; return this; }
        /// <summary>Sets extended property Manager.</summary>
        public InfoBuilder Manager(string manager) { _doc.ApplicationProperties.Manager = manager; return this; }
        /// <summary>Sets extended property Application name.</summary>
        public InfoBuilder Application(string app) { _doc.ApplicationProperties.ApplicationName = app; return this; }
        /// <summary>Sets core property LastModifiedBy.</summary>
        public InfoBuilder LastModifiedBy(string user) { _doc.BuiltinDocumentProperties.LastModifiedBy = user; return this; }
    }
}
