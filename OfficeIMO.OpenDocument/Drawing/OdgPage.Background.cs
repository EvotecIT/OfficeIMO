namespace OfficeIMO.OpenDocument;

/// <summary>Native Draw background paint area.</summary>
public enum OdgBackgroundSize {
    /// <summary>Paint inside the page layout's margins.</summary>
    Border,
    /// <summary>Paint the entire page.</summary>
    Full
}

public sealed partial class OdgPage {
    /// <summary>Page-local background declarations. Page no-fill retains the master's fill in native Draw.</summary>
    /// <remarks>Edits use copy-on-write for shared drawing-page styles. They do not change master paint or artwork.</remarks>
    public OdgBackground Background => new OdgBackground(_document, Element, "content.xml");

    /// <summary>Background declarations of the referenced master. Edits affect every page using this master.</summary>
    /// <remarks>Clone the master and assign its name to <see cref="MasterPageName"/> before independent master edits.</remarks>
    public OdgBackground MasterBackground => new OdgBackground(_document,
        Master ?? throw new InvalidDataException("Drawing page has no resolvable master."), "styles.xml");

    /// <summary>The master's full-page or inside-margin paint area, shared by pages referencing that master.</summary>
    /// <remarks>The native default is <see cref="OdgBackgroundSize.Border"/>. A page-local fill uses this same area.</remarks>
    public OdgBackgroundSize MasterBackgroundSize {
        get => MasterBackground.ReadProperty(OdfNamespaces.Draw + "background-size") switch {
            null or "border" => OdgBackgroundSize.Border,
            "full" => OdgBackgroundSize.Full,
            _ => throw new InvalidDataException("Unknown native background size.")
        };
        set {
            if (value is not (OdgBackgroundSize.Border or OdgBackgroundSize.Full)) throw new ArgumentOutOfRangeException(nameof(value));
            MasterBackground.WriteProperty(OdfNamespaces.Draw + "background-size", value == OdgBackgroundSize.Full ? "full" : "border");
        }
    }
}
