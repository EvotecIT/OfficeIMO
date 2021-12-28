using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Tests
{
    public static class TestHelper
    {
        static TestHelper()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            DirectoryWithFiles = Path.Combine(baseDirectory, "documents");
        }

        public static string DirectoryWithFiles { get; }
    }
}
