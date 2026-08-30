namespace OfficeIMO.Drawing;

/// <summary>Identifies the device class declared by an ICC profile header.</summary>
public enum OfficeIccProfileClass {
    /// <summary>Input-device profile (<c>scnr</c>).</summary>
    InputDevice,

    /// <summary>Display-device profile (<c>mntr</c>).</summary>
    DisplayDevice,

    /// <summary>Output-device profile (<c>prtr</c>).</summary>
    OutputDevice
}
