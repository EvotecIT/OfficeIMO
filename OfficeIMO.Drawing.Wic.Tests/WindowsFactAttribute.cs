using System;
using OfficeIMO.Drawing.Wic;
using Xunit;

namespace OfficeIMO.Drawing.Wic.Tests;

/// <summary>Records a real test skip when WIC is not a supported platform.</summary>
public sealed class WindowsFactAttribute : FactAttribute {
    public WindowsFactAttribute() {
        if (!OfficeWicCodec.IsSupported) Skip = "WIC integration requires Windows.";
    }
}

/// <summary>Records missing user-installed extensions as a skip, rather than successful decode proof.</summary>
public sealed class WindowsHeicFactAttribute : FactAttribute {
    public WindowsHeicFactAttribute() {
        if (!OfficeWicCodec.IsSupported) {
            Skip = "WIC integration requires Windows.";
        } else {
            OfficeRasterCodecAvailability availability = new OfficeWicCodec().GetAvailability();
            if (!availability.IsAvailable) Skip = availability.Message;
        }
    }
}
