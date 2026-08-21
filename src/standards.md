# West Fraser PLC Tag Standards

> Source of truth: the West Fraser **Process Data Collection Standard Procedure**.
> The by-example patterns below were derived from real mill workbooks (Angelina,
> Joyce) to capture conventions the procedure states but does not spell out.
> This file is loaded into the LLM system prompt to guide tag suggestions.
> Suggestions are advisory only — a human reviews and applies them.

---

## 1. Tag Name conventions

A StandardTagName must be:

- **TitleCase**, with **no spaces, underscores, or special characters**.
  - `Canterrunning` → `CanterRunning` (not TitleCase)
  - `Blender_Speed` → `BlenderSpeed` (underscore)
  - `TrimLoss&` → `TrimLoss` (special character)
- **Left / Right** in the machine direction — never East, West, North, South.
  - `EastDryer` → `LeftDryer`
  - `MatHeightEast, MatHeightWest` → `MatHeightInchLeft, MatHeightInchRight`
- **Delineation (numbering) at the END** of the name.
  - `Former1Height` → `FormerHeight1`
  - Real examples: `CenterFan1`, `CenterFan2` … `CenterFan12`
- **Descriptive but succinct** — the qualifier goes where it reads naturally, not
  crammed at the front.
  - `CL2BlenderSpeed` → `BlenderSpeedHzCL2`
  - `BSLFormingBunkerLevel` → `FormingBunkerLevelBSL`

### Naming by example (real PLC tag → StandardTagName)

The StandardTagName describes the *meaning* of the value, not the raw PLC address.
The PLC tag is often cryptic; the StandardTagName is human-readable.

**The StandardTagName depends on the table/area, not just the PLC tag.** The same raw
PLC tag can map to different StandardTagNames in different tables — e.g.
`d_TotalPieces_Out` becomes `BoardCount` in the edger, `PiecesOutput` in the green
trimmer, and `RemanCount` where boards are routed to reman. Always consider which table
the row belongs to when naming.

| PLC tag (raw)                     | StandardTagName    |
| --------------------------------- | ------------------ |
| `d_FullBins`                      | `BinCountFull`     |
| `d_SpareBins`                     | `BinCountEmpty`    |
| `LugS.Speed`                      | `LineSpeedFpm`     |
| `ET_LugSpeed`                     | `LineSpeed`        |
| `d_LugFillPercent`                | `LugFillPct`       |
| `d_TotalPieces_Out`               | `BoardCount` / `PiecesOutput` (context-dependent) |
| `d_FBM_Out`                       | `BoardFoot` / `BoardFootage` |
| `SorterDrive:I.OutputCurrent`     | `SorterChainAmps`  |
| `Mach.Mtr1.MotorCurrent.Current`  | `SawArborAmps`     |
| `CurrentShift`                    | `Shift`            |
| `UptimePercent`                   | `ChainUptime`      |

Common abbreviations seen in names: `Pct` (percent), `Fpm` (feet per minute),
`Hz` (hertz), `Amps`, `BSL`, `CL2` (line/side qualifiers appended at the end).

---

## 2. Data Types

Permitted types actually used, in order of frequency: **float, int, nvarchar(255),
datetime**. (`uniqueidentifier` and `nvarchar(50)` appear only on system/metadata
columns.)

- **BIT is NOT permitted.** BITs cause conflicting display patterns in downstream
  applications. **Use INT instead** — it shows the same 0/1 information but displays
  as a plain number.
  - Note: a few legacy `bit` tags exist in the workbooks (e.g. the Planer "…ZoneRunning"
    tags). This is background context only — do NOT raise BIT as a flag. If a row is
    already flagged for another reason and happens to be BIT, you may mention INT in the
    suggestion, but BIT alone is never a reason to flag a row.

### Which type for which kind of tag (by example)

- **Running / Enabled / Status / InProgress / Active flags → `int`**
  (`CanterRunning`, `PusherEnabled`, `BreakInProgress`, `SorterRunning`,
  `KilnRunning` are all `int`). This is the INT-instead-of-BIT rule in practice.
- **Speeds, currents/amps, volumes, percentages, footage, temperatures → `float`**
  (`CanterLineSpeed`, `SorterChainAmps`, `LugFillPct`, `BoardFootage`).
- **Counts (pieces, boards, bins, logs) → `int`**
  (`BoardCount`, `BinCountFull`, `LineLogsLoaded`, `PieceCount`).
- **Identifiers / text (shift, crew, product, classification) → `nvarchar(255)`**
  (`Shift`, `Crew`, `Product`, `ActiveProduct`, `BoardClassification`).

A one-word exception seen in the data: `AirCompressorStatus` is `float` — status-style
names are usually `int`, so a status/running name typed as `float` is worth a look.

---

## 3. Request only RAW, real-time tags

Only request the **raw, real-time value** of a process variable. Do **not** request
calculated, aggregated, or time-based "snapshot" tags — the historian timestamps every
row, so aggregates can be computed on demand in SQL Views or dataPARC.

- `PlanerLugFillPctLastHour` → request `PlanerLugFillPct`
- `Baghouse1AbortTimePreviousMonth` → request `Baghouse1AbortTime`
- `PressloadsLastMonth` → request `PressloadCount` (the raw event)

Red-flag suffixes that usually indicate a disallowed aggregate: `LastHour`,
`LastMonth`, `PreviousMonth`, `Previous…`, `…Average` (when it's a stored rollup
rather than a live PLC value), `Total…LastX`.

---

## 4. The other request fields

- **PLC IP** and **PLC address (path)** — the physical source of the value.
- **Data Type** — see section 2.
- **Units of measure** — stored in dataPARC so users understand the value. If a needed
  unit isn't in the form dropdown, leave blank and notify ROC/TagsRUs.
- **Transaction frequency** — scheduled (e.g. 5 s) or unscheduled event. A **1 s**
  frequency **requires additional justification** (SQL load/volume).
- **Tag description (TagPurpose)** — shown beside the tag in dataPARC. Keep it short,
  concrete, and about *what the value means*. Match the workbook voice:
  - "Indicates sorter operational state"
  - "Electrical load on the primary sorter chain"
  - "Count of full sorter bins"
  - "Sorter line speed in feet per minute"
- **Min / Max** — default y-axis range in dataPARC only; does **not** limit collected
  data and can change anytime. Provide a reasonable range.
- **NewTable?** — mark tags that belong to a brand-new table. If unsure whether a table
  already exists, leave unchecked and notify ROC/TagsRUs.

---

## 5. Table naming (observed)

Tables are TitleCase and end in `DataTable` for the main per-area tables
(`CanterDataTable`, `EdgerDataTable`, `GreenSorterDataTable`, `PlanerDataTable`,
`KilnDataTable`, `EnvironmentalDataTable`). A few specialized tables use a suffix
instead (`GreenSorterBins`, `DrySorterBins`). Numbered lines append the number at the
end (`Kiln2DataTable`).

---

## 6. What NOT to change

- Do not invent units, frequencies, or Min/Max the requester didn't provide — flag them
  as needing input instead.
- Do not "fix" `DefaultSource` or `Criticality` being blank — those are optional.
- Preserve the existing **Status** value on any suggested row.
- When unsure, say so and lower the confidence rather than guessing a specific value.