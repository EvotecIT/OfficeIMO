using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Provides test setup for Excel documents.
    /// </summary>
    public partial class Excel {
        private readonly string _directoryDocuments;
        private readonly string _directoryWithFiles;
        private readonly string _directoryWithImages;

        public Excel() {
            _directoryDocuments = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            _directoryWithImages = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            //_directoryDocuments = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Tests", "TempDocuments");
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments1");
            Word.Setup(_directoryWithFiles);
        }
    }
}