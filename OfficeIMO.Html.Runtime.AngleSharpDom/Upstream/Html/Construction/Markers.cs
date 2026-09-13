namespace AngleSharp.Html.Construction;

using System;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// Represents a constructable template element.
/// </summary>
public interface IConstructableTemplateElement : IConstructableElement
{
    /// <summary>
    /// Populates the fragment with the content of the template.
    /// </summary>
    void PopulateFragment();
}

/// <summary>
/// Represents a constructable SVG element.
/// </summary>
public interface IConstructableSvgElement : IConstructableElement;

/// <summary>
/// Represents a constructable script element.
/// </summary>
public interface IConstructableScriptElement: IConstructableElement
{
    /// <summary>
    /// Runs a prepared inline non-module script on the current script stack.
    /// </summary>
    /// <returns>True when synchronous execution was available.</returns>
    Boolean RunSynchronously();

    /// <summary>
    ///
    /// </summary>
    /// <param name="cancel"></param>
    /// <returns></returns>
    Task RunAsync(CancellationToken cancel);

    /// <summary>
    ///
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    Boolean Prepare(IConstructableDocument document);
}

/// <summary>
/// Marks a style or stylesheet link created by the active document parser.
/// </summary>
public interface IConstructableStyleSheetElement : IConstructableElement
{
    /// <summary>
    /// Records parser creation before the element is initialized.
    /// </summary>
    void MarkParserInserted();
}

/// <summary>
/// Represents a constructable meta element.
/// </summary>
public interface IConstructableMetaElement : IConstructableElement
{
    /// <summary>
    /// Handles the meta element.
    /// </summary>
    void Handle();
}

/// <summary>
/// Represents a constructable math element.
/// </summary>
public interface IConstructableMathElement : IConstructableElement;

/// <summary>
/// Represents a constructable frame element.
/// </summary>
public interface IConstructableFrameElement : IConstructableElement;

/// <summary>
/// Represents a constructable form element.
/// </summary>
public interface IConstructableFormElement : IConstructableElement;
