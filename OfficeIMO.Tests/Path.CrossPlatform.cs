using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public class PathCrossPlatform {
        [Fact]
        public void PathCombine_AppendsSeparator_ForDirectoryUri() {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            try {
                var baseHref = new Uri(new Uri(Path.Combine(dir, "dummy"), UriKind.Absolute), ".").AbsoluteUri;
                Assert.EndsWith("/", baseHref);
            } finally {
                Directory.Delete(dir);
            }
        }
    }
}

