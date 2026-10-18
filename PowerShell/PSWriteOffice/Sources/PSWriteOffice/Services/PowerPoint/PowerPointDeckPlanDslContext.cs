using System;
using System.Threading;
using OfficeIMO.PowerPoint;
using System.Management.Automation;

namespace PSWriteOffice.Services.PowerPoint;

internal sealed class PowerPointDeckPlanDslContext : IDisposable
{
    private static readonly AsyncLocal<PowerPointDeckPlanDslContext?> CurrentScope = new();

    private PowerPointDeckPlanDslContext(PowerPointDeckPlan plan)
    {
        Plan = plan ?? throw new ArgumentNullException(nameof(plan));
    }

    public PowerPointDeckPlan Plan { get; }

    public static PowerPointDeckPlanDslContext? Current => CurrentScope.Value;

    public static PowerPointDeckPlanDslContext Enter(PowerPointDeckPlan plan)
    {
        if (plan == null)
        {
            throw new ArgumentNullException(nameof(plan));
        }

        if (CurrentScope.Value != null)
        {
            throw new InvalidOperationException("A PowerPoint deck-plan DSL scope is already active on this runspace.");
        }

        var scope = new PowerPointDeckPlanDslContext(plan);
        CurrentScope.Value = scope;
        return scope;
    }

    public static PowerPointDeckPlanDslContext Require(PSCmdlet caller)
    {
        var scope = CurrentScope.Value;
        if (scope == null)
        {
            throw new InvalidOperationException(
                $"'{caller.MyInvocation.InvocationName}' must run inside New-OfficePowerPointDeckPlan.");
        }

        return scope;
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = null;
        }
    }
}
