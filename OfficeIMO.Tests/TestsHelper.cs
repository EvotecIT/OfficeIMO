using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Tests
{
    public static class TestsHelper
    {
        public static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            } else {
                Directory.Delete(path, true);
                Directory.CreateDirectory(path);
            }
        }
        static TestsHelper()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            DirectoryWithFiles = Path.Combine(baseDirectory, "documents");
        }

        public static string DirectoryWithFiles { get; }
    }
}
