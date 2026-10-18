using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

internal sealed class PowerPointDslContext : IDisposable
{
    private static readonly AsyncLocal<PowerPointDslContext?> CurrentScope = new();
    private readonly Stack<PowerPointSlide> _slides = new();

    private PowerPointDslContext(PowerPointPresentation presentation)
    {
        Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
    }

    public PowerPointPresentation Presentation { get; }

    public static PowerPointDslContext? Current => CurrentScope.Value;

    public static PowerPointDslContext Enter(PowerPointPresentation presentation)
    {
        if (presentation == null)
        {
            throw new ArgumentNullException(nameof(presentation));
        }

        if (CurrentScope.Value != null)
        {
            throw new InvalidOperationException("A PowerPoint DSL scope is already active on this runspace.");
        }

        var scope = new PowerPointDslContext(presentation);
        CurrentScope.Value = scope;
        return scope;
    }

    public static PowerPointDslContext Require(PSCmdlet caller)
    {
        var scope = CurrentScope.Value;
        if (scope == null)
        {
            throw new InvalidOperationException(
                $"'{caller.MyInvocation.InvocationName}' must run inside New-OfficePowerPoint.");
        }

        return scope;
    }

    public PowerPointSlide? CurrentSlide => _slides.LastOrDefault();

    public PowerPointSlide RequireSlide()
    {
        var slide = CurrentSlide;
        if (slide == null)
        {
            throw new InvalidOperationException("No slide context available. Use Add-OfficePowerPointSlide / PptSlide first.");
        }

        return slide;
    }

    public IDisposable Push(PowerPointSlide slide)
    {
        if (slide == null)
        {
            throw new ArgumentNullException(nameof(slide));
        }

        _slides.Push(slide);
        return new PopToken(this, slide);
    }

    private void Pop(PowerPointSlide slide)
    {
        if (_slides.Count == 0)
        {
            return;
        }

        if (ReferenceEquals(_slides.Peek(), slide))
        {
            _slides.Pop();
        }
    }

    private sealed class PopToken : IDisposable
    {
        private PowerPointDslContext? _context;
        private readonly PowerPointSlide _slide;

        public PopToken(PowerPointDslContext context, PowerPointSlide slide)
        {
            _context = context;
            _slide = slide;
        }

        public void Dispose()
        {
            _context?.Pop(_slide);
            _context = null;
        }
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = null;
        }

        _slides.Clear();
    }
}
