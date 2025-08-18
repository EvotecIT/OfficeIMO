using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio document containing pages.
    /// </summary>
    public class VisioDocument {
        private readonly List<VisioPage> _pages = new();

        /// <summary>
        /// Collection of pages in the document.
        /// </summary>
        public IReadOnlyList<VisioPage> Pages => _pages;

        /// <summary>
        /// Adds a new page to the document.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        public VisioPage AddPage(string name) {
            VisioPage page = new(name);
            _pages.Add(page);
            return page;
        }
    }
}

