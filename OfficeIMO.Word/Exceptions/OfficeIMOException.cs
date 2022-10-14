using System;

namespace OfficeIMO.Word;

public abstract class OfficeIMOException : Exception
{
    protected OfficeIMOException(string message) : base(message) { }
}
