using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private readonly string _directoryDocuments;
        private readonly string _directoryWithFiles;
        private readonly string _directoryWithImages;

        internal static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            } else {
                //Directory.Delete(path, true);
                //Directory.CreateDirectory(path);
            }
        }

        public Word() {
            _directoryDocuments = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            _directoryWithImages = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            //_directoryDocuments = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Tests", "TempDocuments");
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments2");
            Setup(_directoryWithFiles);
        }

        /// <summary>
        /// This helps finding unexpected elements during validation. Should prevent unexpected changes
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public bool HasUnexpectedElements(WordDocument document) {
            bool found = false;
            foreach (var e in document.DocumentValidationErrors) {
                if (e.Description.StartsWith("The element has unexpected child element")) {
                    found = true;
                    break;
                }
            }
            return found;
        }
    }
}
