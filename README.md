# Excel Script Player

A lightweight script player engine built with Excel VBA.

Excel Script Player reads a text-based script, parses its control information, and plays the script according to its timing information.

It is designed as a small **Player / Engine experiment in Excel VBA** rather than as a game or novel-making tool.

---

## Overview

Excel Script Player follows a simple processing flow:

```text
Script
  ↓
Parser
  ↓
Internal Data
  ↓
Player / Engine
  ↓
Display
```

The script is stored as a plain text file.

The Player reads the script, identifies timing information and control tags, and displays the text according to the selected playback mode.

The current display is implemented with an Excel UserForm.

The same basic Player / Engine concept could potentially be connected to other display systems in the future.

---

## Requirements

The current project is developed and tested with:

* Microsoft Excel 2007
* VBA 6.5
* Windows 7
* 32-bit environment

The project uses classic VBA / UserForm features and is intentionally kept compatible with older Excel environments.

---

## How to Run

1. Download or clone this repository.
2. Keep `ExcelPlayer.xlsm` and the `Script` folder in the same location.
3. Open `ExcelPlayer.xlsm`.
4. The Player reads:

```text
Script/EndRoll.txt
```

5. Run the Player.

The repository already contains an example `EndRoll.txt`, so the Player can be tested immediately.

---

## Run and Preview

Excel Script Player provides two ways to start the Player.

### Run

`Run` starts the script playback directly.

The script is loaded and played using the current Player settings.

This mode is intended for normal playback.

### Preview

`Preview` checks the script before starting playback.

The script is loaded, its script range is determined, and the supported script structures, tags, time information, and timeline settings are validated before the UserForm is displayed.

This mode is useful when creating or editing a script, because script errors can be detected before playback.

The current validation includes:

* `<SCRIPT>` / `</SCRIPT>` structure
* `<TIME>` / `</TIME>` structure
* Tag structure
* Time counter format
* Time information
* Timeline settings

In short:

```text
Run
  → Load
  → Play

Preview
  → Load
  → Determine Script Range
  → Validate
  → Prepare Time Data
  → Play
```

---

## Sample Data

The `SampleData` folder contains small scripts for testing different playback features.

```text
SampleData/
├─ 01_Basic.txt
├─ 02_TimeCounter.txt
├─ 03_TimeTag.txt
└─ 04_Wait.txt
```

### 01_Basic.txt

A basic script without a time counter.

It is useful for checking the basic script loading and display process.

### 02_TimeCounter.txt

A script using a line-leading time counter.

Example:

```text
00:01:50:私はExcel Script Player
00:01:58:最近ちょっと生まれ変わった
```

### 03_TimeTag.txt

A script using `<TIME>...</TIME>` tags.

Example:

```text
<TIME>00:01:50:</TIME>私はExcel Script Player
<TIME>00:01:58:</TIME>最近ちょっと生まれ変わった
```

### 04_Wait.txt

A script designed for `TIME_WAIT` playback.

The time information is used as a wait interval between displayed lines.

---

## Script Format

A basic script is enclosed by:

```text
<SCRIPT>

script contents

</SCRIPT>
```

The `<SCRIPT>` and `</SCRIPT>` tags define the script range.

The tags themselves are not displayed.

Timing information can be specified either as a line-leading time counter or with `<TIME>` tags.

### Line-leading time counter

Two formats are supported:

```text
00:00:00Hello
```

and

```text
00:00:00:Hello
```

The Player can use these values for timeline playback or wait-based playback.

### `<TIME>` tag

Time information can also be written separately from the displayed text:

```text
<TIME>00:00:05:</TIME>Hello
```

The display of the time text can be controlled independently from its use as timing information.

---

## Playback Modes

The Player currently supports three timing modes:

```text
TIME_IGNORE
TIME_WAIT
TIME_TIMELINE
```

### TIME_IGNORE

Timing information is ignored.

Lines are displayed without using the time information for playback control.

### TIME_WAIT

The time value is used as a wait interval.

For example:

```text
<TIME>00:00:05:</TIME>First line
<TIME>00:00:08:</TIME>Second line
```

can be used to control the waiting time for each line.

A default wait time can also be used when no timing information is present.

### TIME_TIMELINE

The time values are treated as positions on a timeline.

The Player waits until the current playback time reaches the specified position before displaying the corresponding line.

The first time value can be interpreted in different ways, including:

* First value as the timeline start
* First value as zero
* Start offset

---

## Display Modes

The Player currently provides two display modes:

```text
DISPLAY_MESSAGE
DISPLAY_ENDROLL
```

`DISPLAY_MESSAGE` displays text from the configured starting position.

`DISPLAY_ENDROLL` starts the text from the bottom of the display and scrolls it upward, similar to a movie end roll.

The display itself is currently implemented using dynamically created MSForms Label controls on a UserForm.

---

## Player Settings

The default Player settings are defined in `Player_Init`.

The following settings are used as the initial configuration:

```vb
Public Sub Player_Init()

    RepeatWaitSeconds = 8

    '----------------------------------
    ' Display settings
    '----------------------------------
    ShowTimeCounter = True
    ShowTime = True

    ' DISPLAY_MESSAGE or DISPLAY_ENDROLL
    DisplayMode = DISPLAY_MESSAGE
    FirstLineRows = 1

    '----------------------------------
    ' Time settings
    '----------------------------------

    ' TIME_TIMELINE or TIME_WAIT
    TimeCounterMode = TIME_TIMELINE
    DefaultWaitSeconds = 0.6

    TimeLineMode = _
        TIMELINE_FIRST_AS_ZERO

    '----------------------------------
    ' TimeCounter settings
    '----------------------------------
    TimeCounterFormat = _
        TIME_COUNTER_FORMAT_9

    '----------------------------------
    ' Repeat settings
    '----------------------------------
    ScriptRepeatMode = False

End Sub
```

These settings can be changed to experiment with different playback behaviors.

### Display

`ShowTimeCounter`

Controls whether a line-leading time counter is displayed.

```text
True  → 00:00:05:Hello
False → Hello
```

`ShowTime`

Controls whether the value inside `<TIME>...</TIME>` is displayed.

`DisplayMode` selects the display style:

```text
DISPLAY_MESSAGE
DISPLAY_ENDROLL
```

`FirstLineRows` specifies the logical starting row for message display.

### Time

`TimeCounterMode` selects how time information is used:

```text
TIME_IGNORE
TIME_WAIT
TIME_TIMELINE
```

`DefaultWaitSeconds` specifies the default wait time when `TIME_WAIT` is used without timing information.

`TimeLineMode` controls how the first timeline value is interpreted.

Available modes include:

```text
TIMELINE_FIRST_AS_START
TIMELINE_FIRST_AS_ZERO
TIMELINE_START_OFFSET
```

### TimeCounter Format

Two display formats are supported:

```text
TIME_COUNTER_FORMAT_8

00:00:05Hello
```

and:

```text
TIME_COUNTER_FORMAT_9

00:00:05:Hello
```

### Repeat

`ScriptRepeatMode` controls whether the script is repeated.

```text
False → stop after playback
True  → repeat the script
```

`RepeatWaitSeconds` specifies the waiting time before a repeated playback starts.

---

## Validation

The project includes a script validation system.

The current validation scope includes:

* `<SCRIPT>` / `</SCRIPT>` structure
* `<TIME>` / `</TIME>` structure
* Time format validation

Validation is intended to detect script structure errors before playback.

Additional script tags may be added in future versions.

---

## Project Structure

```text
excel-script-player/
│
├─ ExcelPlayer.xlsm
│
├─ Script/
│   └─ EndRoll.txt
│
├─ SampleData/
│   ├─ 01_Basic.txt
│   ├─ 02_TimeCounter.txt
│   ├─ 03_TimeTag.txt
│   └─ 04_Wait.txt
│
└─ src/
    ├─ PLY_01_Config.bas
    ├─ PLY_03_Init.bas
    ├─ PLY_04_Script.bas
    ├─ PLY_05_Show.bas
    ├─ PLY_10_Tag.bas
    ├─ PLY_20_Dictionary.bas
    ├─ PLY_30_Validate.bas
    ├─ PLY_40_Timer.bas
    ├─ UserForm1.frm
    └─ UserForm1.frx
```

### Source modules

| Module              | Responsibility                           |
| ------------------- | ---------------------------------------- |
| `PLY_01_Config`     | Configuration and public settings        |
| `PLY_03_Init`       | Player initialization                    |
| `PLY_04_Script`     | Script loading and script range handling |
| `PLY_05_Show`       | Text display and scrolling               |
| `PLY_10_Tag`        | Script tag and display parsing           |
| `PLY_20_Dictionary` | Time information management              |
| `PLY_30_Validate`   | Script validation                        |
| `PLY_40_Timer`      | Timing and playback control              |
| `UserForm1`         | Current UserForm display                 |

The source files in `src` are exported VBA components from the Excel workbook.

---

## Why Excel?

This project is not intended to turn Excel into a full-scale application platform.

Instead, Excel is used as a convenient environment for experimenting with:

* Script parsing
* Timing control
* Playback engines
* State and data handling
* Display systems
* Small reusable engines

Excel provides a familiar environment where the script, engine, and display can be tested together.

---

## Project Status

This is an experimental Player / Engine project.

The current version focuses on establishing the basic architecture:

```text
Script
→ Parse
→ Store
→ Control Time
→ Play
→ Display
```

More script controls and display features may be added in future versions.

---

## License

This project is released under the MIT License.

See [LICENSE](LICENSE) for details.
