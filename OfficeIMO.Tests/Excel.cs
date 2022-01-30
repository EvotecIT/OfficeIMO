using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static string _directoryDocuments;
        private readonly string _directoryWithFiles;

        public Excel() {
            _directoryDocuments = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Tests", "TempExcelDocuments");
            TestsHelper.Setup(_directoryDocuments); // prepare temp documents directory 
            _directoryWithFiles = TestsHelper.DirectoryWithFiles;
        }
    }
}