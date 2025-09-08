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

        public InfoBuilder Title(string title) { _doc.BuiltinDocumentProperties.Title = title; return this; }
        public InfoBuilder Author(string author) { _doc.BuiltinDocumentProperties.Creator = author; return this; }
        public InfoBuilder Subject(string subject) { _doc.BuiltinDocumentProperties.Subject = subject; return this; }
        public InfoBuilder Keywords(string keywords) { _doc.BuiltinDocumentProperties.Keywords = keywords; return this; }
        public InfoBuilder Description(string description) { _doc.BuiltinDocumentProperties.Description = description; return this; }
        public InfoBuilder Category(string category) { _doc.BuiltinDocumentProperties.Category = category; return this; }
        public InfoBuilder Company(string company) { _doc.ApplicationProperties.Company = company; return this; }
        public InfoBuilder Manager(string manager) { _doc.ApplicationProperties.Manager = manager; return this; }
        public InfoBuilder Application(string app) { _doc.ApplicationProperties.ApplicationName = app; return this; }
        public InfoBuilder LastModifiedBy(string user) { _doc.BuiltinDocumentProperties.LastModifiedBy = user; return this; }
    }
}

