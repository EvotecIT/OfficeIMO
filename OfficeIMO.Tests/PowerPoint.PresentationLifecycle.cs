using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPresentationLifecycle {
        [Fact]
        public void SaveThrowsAfterDispose() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);

            presentation.Dispose();

            Assert.Throws<ObjectDisposedException>(() => presentation.Save());

            File.Delete(filePath);
        }
    }
}
