Absolutely—build a tiny **“Visio Inspector”** + **typed shape profiles** pipeline. It lets you load any vendor stencil (AWS/Azure/Cisco, BPMN, etc.), **normalize** it to a stable snapshot, **diff** against your output, and then **generate or refine typed shape support**. Below is a concrete plan with APIs and usage patterns—no implementation details, just how it should look and behave.

---

# Visio Shape Typing + Diffing — Design

## Goals

* **Understand** unknown shapes/stencils (masters) by inspecting live `.vsdx`.
* **Normalize** noisy XML → stable snapshots for comparison.
* **Diff** “expected” vs “ours” to see what we’re missing.
* **Type** common shapes via small, composable “profiles” (e.g., Flowchart.Process, AWS.EC2).
* **Generate** shims/templates for new shapes with minimal manual work.

---

## 1) Snapshot Model (stable, testable)

A canonical POCO you can serialize to JSON/YAML for golden tests.

```csharp
public sealed class VisioSnapshot {
  public string? Title { get; init; }
  public string? Author { get; init; }
  public List<VisioPageSnap> Pages { get; init; } = new();
}

public sealed class VisioPageSnap {
  public string Name { get; init; } = "";
  public double WidthIn { get; init; }
  public double HeightIn { get; init; }
  public List<VisioShapeSnap> Shapes { get; init; } = new();
  public List<VisioConnectorSnap> Connectors { get; init; } = new();
}

public sealed class VisioShapeSnap {
  public string Id { get; init; } = "";               // e.g., "Task1"
  public string? MasterName { get; init; }            // e.g., "Process"
  public string? StencilName { get; init; }           // e.g., "Flowchart.vssx"
  public string? Text { get; init; }
  public double X { get; init; }                      // inches, normalized
  public double Y { get; init; }
  public double W { get; init; }
  public double H { get; init; }
  public Dictionary<string,string> Data { get; init; } = new(); // Shape Data / Props
  public Dictionary<string,string> Style { get; init; } = new(); // Fill, Line, etc.
}

public sealed class VisioConnectorSnap {
  public string Id { get; init; } = "";               // e.g., "Path1"
  public string From { get; init; } = "";             // shape id
  public string To { get; init; } = "";               // shape id
  public string Kind { get; init; } = "RightAngle";   // normalized
  public string? Label { get; init; }
  public string? ArrowBegin { get; init; }
  public string? ArrowEnd { get; init; }
}
```

### Usage

```csharp
var snap = VisioInspector.Snapshot("aws-example.vsdx",
    new SnapshotOptions { NormalizePositions = true, RoundDigits = 3 });

File.WriteAllText("aws-example.snapshot.json", snap.ToJson());
```

---

## 2) Normalization (kill noise, keep meaning)

`SnapshotOptions`:

* `NormalizePositions` (snap to grid; round to 1/8 in)
* `RoundDigits` (e.g., 3 decimals)
* `IgnoreGuidLikeIds` (map transient ids to stable aliases)
* `IncludeStyles` (true/false)
* `IncludeData` (true/false)
* `SortBy` (e.g., “by X then Y then Name”)

Normalization ensures two visually-identical diagrams compare equal even if raw XML varies.

---

## 3) Differ (what’s different and where)

A human-readable structural diff that highlights missing types or mapping gaps.

```csharp
public sealed class VisioDiff {
  public bool IsEqual { get; init; }
  public List<string> Messages { get; init; } = new(); // human-friendly lines
  public VisioSnapshot? LeftOnly { get; init; }
  public VisioSnapshot? RightOnly { get; init; }
}

var expected = VisioInspector.Snapshot("expected.vsdx");
var actual   = VisioInspector.Snapshot("ours.vsdx");

var diff = VisioInspector.Diff(expected, actual,
    new DiffOptions { ToleranceInches = 0.05, CompareStyles = false });

Console.WriteLine(diff.ToText());
```

**Examples of messages**

* `Missing shape ‘Decision1’ (master=Decision) on page ‘Process’`
* `Connector ‘PathYes’ kind differs: expected RightAngle, got Straight`
* `Shape ‘EC2’: size differs beyond tolerance (W: expected 1.00, got 0.90)`

---

## 4) Typed Shape Profiles (the “types” you want)

A compact registry that knows how to **recognize** and **emit** common shapes.

```csharp
public interface IShapeProfile {
  string Name { get; }                         // "Flowchart.Process", "AWS.EC2"
  bool Match(VisioShapeSnap snap);             // recognition (from snapshot)
  void Emit(ShapeBuilder s);                   // emission (to fluent builder)
  IReadOnlyDictionary<string,string> RequiredData { get; } // e.g., "Owner"
}

public static class ShapeProfiles {
  public static readonly IShapeProfile Flow_Process = new FlowProcessProfile();
  public static readonly IShapeProfile Flow_Decision = new FlowDecisionProfile();
  public static readonly IShapeProfile AWS_EC2 = new AwsEc2Profile();
  // ...
}
```

### Usage (recognition)

```csharp
var snap = VisioInspector.Snapshot("vendor.vsdx");
foreach (var sh in snap.Pages[0].Shapes) {
  var profile = ShapeProfiles.All.FirstOrDefault(p => p.Match(sh));
  Console.WriteLine($"{sh.Id}: {(profile?.Name ?? "Unknown")}");
}
```

### Usage (emission)

```csharp
doc.AsFluent().Page("P", p => {
  var s = p.Shape("Task1", _ => {});
  ShapeProfiles.Flow_Process.Emit(s);
  s.At(1,5).Size(2.5,1.0).Text("Validate");
});
```

> Over time you’ll build profiles for Flowchart, BPMN, AWS, Azure, Cisco.
> Each profile includes a **Match** predicate (MasterName pattern, geometry hints, default styles) and an **Emit** recipe.

---

## 5) Recorder (capture while using Visio UI)

If devs tweak a diagram manually in Visio, you can **record** changes to enrich profiles.

```csharp
VisioRecorder.Begin(doc, new RecorderOptions { TrackMoves = true, TrackStyle = true });
// user makes changes in code or via helpers
VisioRecorder.End(out var ops);

foreach (var op in ops) Console.WriteLine(op); // e.g., "Move Task1: (1.00,5.00)->(1.25,5.00)"
```

This is optional but handy for reverse-engineering vendor stencils into profiles.

---

## 6) “Unknown” shape support (fallback)

Not everything needs a profile. Provide safe fallbacks:

```csharp
page.AddUnknownShape(id: "SvgIcon", x: 3, y: 4, w: 1, h: 1)
    .FromImage("logo.svg")       // or FromPng/FromEmf
    .Text("My Icon");
```

The snapshot/diff tools will still see it; you can promote it to a typed profile later.

---

## 7) CLI / Automation (nice to have)

Add a tiny CLI for CI/PRs:

```bash
officeimo-visio inspect vendor.vsdx --out vendor.snapshot.json
officeimo-visio diff expected.snapshot.json ours.snapshot.json --tolerance 0.05
officeimo-visio recognize vendor.snapshot.json --profiles flow,bpmn,aws
```

Outputs non-zero exit code on hard diffs; perfect for golden-file testing.

---

## 8) Workflow for adding a new typed shape

1. **Collect** vendor example `.vsdx` with the shape.
2. **Snapshot** → JSON.
3. **Recognize** → see if any profile matches; if not, scaffold:

```csharp
var scaffold = VisioInspector.ScaffoldProfileFromSnapshot(snapShape);
File.WriteAllText("AwsAlbProfile.cs", scaffold.Source);
```

4. **Refine** the profile (match heuristics, default style).
5. **Emit** test: generate our diagram with that profile.
6. **Diff** expected vs ours; iterate until equal (within tolerance).
7. **Commit** profile + golden snapshots.

---

## 9) What counts as “connector types” and arrows

* **Connector kinds**: `Straight`, `RightAngle` (orthogonal), `Curved`, `Dynamic` (auto route).
* **Arrowheads**: `None`, `Triangle`, `Stealth`, `Diamond`, `Oval`, `OpenArrow`, `Circle`, `Square` … (\~15 common).
  Expose as enums; map to ShapeSheet line ends internally. Snapshot stores normalized names.

---

## 10) API Surface Summary (what you’ll ship)

### Inspector / Diff

```csharp
public static class VisioInspector {
  public static VisioSnapshot Snapshot(string vsdxPath, SnapshotOptions? o = null);
  public static VisioDiff Diff(VisioSnapshot expected, VisioSnapshot actual, DiffOptions? o = null);
  public static string ScaffoldProfileFromSnapshot(VisioShapeSnap snap); // generates starter code
}
```

### Fluent + Profiles

```csharp
// Stencils & profiles
doc.AsFluent()
  .Stencil(st => st.Use(BuiltInStencil.Flowchart).UseFile("AWS.vssx"))
  .Page("Arch", p => p
      .Shape("EC2", s => ShapeProfiles.AWS_EC2.Emit(s).At(3,4).Text("Web")))
  .End();
```

### Validation / Snapshot

```csharp
var report = doc.Validate(v => v.RequireUniqueShapeIds().RequireAllConnectorsResolved());
var snap = doc.Snapshot(page: "Arch"); // same model as Inspector
```

---

## 11) Naming & DX tips

* Keep **IDs human-friendly**; provide `Names.Safe()` & uniqueness guard.
* Make **units explicit** (inches/cm) in API names: `.AtInches(x,y)` vs `.At(x,y)` if you need both.
* Prefer **property-like verbs** in fluent (`Size`, `At`, `Text`, `Data`, `Theme`), no `SetXxx`.
* Offer **presets** for popular families (Flowchart, AWS, Azure) but always allow a **raw Master(name)** path.

---

### Bottom line

* Build a **snapshot/diff** tool once—use it everywhere (profiles, tests, vendor parity).
* Define a **typed profile** interface so shapes are discoverable and generatable.
* Keep **fallbacks** for unknown shapes so users aren’t blocked.
* This gives you a maintainable path to first-class support for **AWS/Azure/Cisco**, BPMN/UML, and whatever comes next—without drowning in raw XML.
