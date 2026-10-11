namespace OfficeIMO.Access {
    public sealed partial class AccessApplicationObject {
        /// <summary>Creates or replaces code-behind for this existing native form or report. Call Save to persist it.</summary>
        /// <remarks>The native class identity remains bound to this object. New class names must be ASCII VBA identifiers of at most 31 characters including Form_ or Report_. Designer and event properties are preserved; source is never executed or compiled. The supplied limits apply to reading the existing project and writing its replacement.</remarks>
        public void SetCodeBehind(string source, OfficeVbaWriteOptions? options = null, CancellationToken cancellationToken = default) {
            EnsureAttached();
            Document.SetCodeBehind(this, source, options, cancellationToken);
        }

        /// <summary>Sets or clears an inert event binding on this existing native form, report or named control.</summary>
        /// <remarks>Null or empty clears the binding. Use [Event Procedure] for existing code-behind; expressions and macro names remain inert. Replacing Click removes its associated embedded macro. New embedded macro definitions are outside this operation.</remarks>
        public void SetEventBinding(AccessEventKind eventKind, string? expression, string? controlName = null,
            OfficeVbaWriteOptions? options = null, CancellationToken cancellationToken = default) {
            EnsureAttached(); Document.SetEventBinding(this, eventKind, expression, controlName, options, cancellationToken);
        }
    }

    /// <summary>Qualified inert form/report and control events.</summary>
    public enum AccessEventKind {
        /// <summary>Form or report Open.</summary>
        Open,
        /// <summary>Named control Click.</summary>
        Click,
        /// <summary>Named control AfterUpdate.</summary>
        AfterUpdate
    }
}
