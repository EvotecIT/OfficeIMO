namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for workbook properties (core + extended).
    /// </summary>
    public sealed class ExcelInfoBuilder {
        private readonly ExcelDocument _doc;
        internal ExcelInfoBuilder(ExcelDocument doc) { _doc = doc; }

        /// <summary>Sets core property Title.</summary>
        public ExcelInfoBuilder Title(string title) { _doc.BuiltinDocumentProperties.Title = title; return this; }
        /// <summary>Sets core property Creator/Author.</summary>
        public ExcelInfoBuilder Author(string author) { _doc.BuiltinDocumentProperties.Creator = author; return this; }
        /// <summary>Sets core property Subject.</summary>
        public ExcelInfoBuilder Subject(string subject) { _doc.BuiltinDocumentProperties.Subject = subject; return this; }
        /// <summary>Sets core property Keywords.</summary>
        public ExcelInfoBuilder Keywords(string keywords) { _doc.BuiltinDocumentProperties.Keywords = keywords; return this; }
        /// <summary>Sets core property Description.</summary>
        public ExcelInfoBuilder Description(string description) { _doc.BuiltinDocumentProperties.Description = description; return this; }
        /// <summary>Sets core property Category.</summary>
        public ExcelInfoBuilder Category(string category) { _doc.BuiltinDocumentProperties.Category = category; return this; }
        /// <summary>Sets extended property Company.</summary>
        public ExcelInfoBuilder Company(string company) { _doc.ApplicationProperties.Company = company; return this; }
        /// <summary>Sets extended property Manager.</summary>
        public ExcelInfoBuilder Manager(string manager) { _doc.ApplicationProperties.Manager = manager; return this; }
        /// <summary>Sets extended property Application name.</summary>
        public ExcelInfoBuilder Application(string app) { _doc.ApplicationProperties.ApplicationName = app; return this; }
        /// <summary>Sets core property LastModifiedBy.</summary>
        public ExcelInfoBuilder LastModifiedBy(string user) { _doc.BuiltinDocumentProperties.LastModifiedBy = user; return this; }
    }
}
