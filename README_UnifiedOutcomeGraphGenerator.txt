# Unified Outcome Graph Generator

An Excel VBA macro for generating Prism-style bar charts with **mean ± SEM** and **individual-value overlays** from one-way or two-way experimental data.

The macro combines three graphing workflows into one guided tool:

- All outcomes on one graph, clustered by experimental group or factor level
- All outcomes on one graph, clustered by outcome
- One separate graph for each outcome

It also centers individual observations over their corresponding bars, offsets points only when values overlap or nearly overlap, and synchronizes the primary and secondary Y-axes using limits calculated from the plotted data.

---

## Contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Running the macro](#running-the-macro)
- [Input data formats](#input-data-formats)
- [Graph options](#graph-options)
- [Individual-value placement](#individual-value-placement)
- [Automatic Y-axis scaling](#automatic-y-axis-scaling)
- [Output worksheets](#output-worksheets)
- [PNG export](#png-export)
- [Statistical calculations](#statistical-calculations)
- [Missing data](#missing-data)
- [Troubleshooting](#troubleshooting)
- [Customization](#customization)
- [Important interpretation notes](#important-interpretation-notes)

---

## Requirements

- Microsoft Excel for Windows or Mac with VBA support
- A workbook saved in the **Excel Macro-Enabled Workbook (`.xlsm`)** format
- Macros enabled for the workbook
- The module file:

```text
UnifiedOutcomeGraphGenerator.bas
```

The macro uses late binding for its internal dictionaries, so you should not need to add a separate VBA reference for `Microsoft Scripting Runtime`.

---

## Installation

1. Open the Excel workbook that will contain the macro.
2. Save it as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Press **Alt+F11** to open the Visual Basic Editor.
   - On some Mac keyboards, use **Fn+Option+F11** or open the VBA editor from the **Developer** tab.
4. In the Visual Basic Editor, select your workbook in the **Project Explorer**.
5. Choose **File → Import File**.
6. Select `UnifiedOutcomeGraphGenerator.bas`.
7. Choose **Debug → Compile VBAProject**.
8. Save the workbook.

If an older copy of the module is already installed, remove it before importing the updated version:

1. Right-click the old `UnifiedOutcomeGraphGenerator` module.
2. Select **Remove UnifiedOutcomeGraphGenerator**.
3. Select **No** when asked whether to export it, unless you want a backup.
4. Import the updated `.bas` file.

---

## Running the macro

1. Arrange the source data in one of the supported formats below.
2. Press **Alt+F8** in Excel.
3. Select:

```text
GenerateUnifiedOutcomeGraphs
```

4. Click **Run**.
5. When prompted, select the complete source table, including the header row.
6. Follow the prompts to choose the experimental design, graph layout, clustering, orientation, and optional PNG export.

The selected source range must be a single contiguous block of cells.

---

## Input data formats

Each row should represent one independent observation or experimental unit. The first row must contain column headers.

### One-way design

Use the first column for the experimental group and all remaining columns for outcomes.

```text
Group | Outcome 1 | Outcome 2 | Outcome 3
A     | 12.4      | 7.1       | 44
A     | 13.2      | 6.8       | 47
B     | 18.5      | 9.4       | 62
B     | 17.9      | 8.8       | 59
```

Expected layout:

```text
Group | Outcome 1 | Outcome 2 | ... | Outcome N
```

### Two-way design

Use the first two columns for the experimental factors and all remaining columns for outcomes.

```text
Factor A | Factor B | Outcome 1 | Outcome 2
Control  | Vehicle  | 12.4      | 7.1
Control  | Drug     | 16.8      | 8.3
Injury   | Vehicle  | 22.1      | 12.7
Injury   | Drug     | 17.5      | 9.8
```

Expected layout:

```text
Factor A | Factor B | Outcome 1 | Outcome 2 | ... | Outcome N
```

### Data-entry recommendations

- Use consistent spelling for group and factor names.
- Avoid leading or trailing spaces in headers and group labels.
- Store outcome measurements as numeric values.
- Leave unavailable measurements blank rather than entering text such as `N/A` where possible.
- Do not include summary rows, subtotals, means, or SEM values in the selected raw-data range.
- Use unique outcome headers.

---

## Graph options

### 1. Experimental design

The first prompt asks whether the data represent:

```text
1 = One-way
2 = Two-way
```

### 2. Graph layout

The next prompt asks whether to create:

```text
1 = All outcomes on ONE graph
2 = One SEPARATE graph per outcome
```

### 3. Clustering for a single combined graph

If all outcomes are placed on one graph, choose one of the following:

#### Groups or factor levels on the X-axis

```text
1 = Experimental groups/factor levels on the X axis;
    outcomes are bar series
```

This is useful when the main comparison is between experimental groups and each outcome should appear as a separate bar series.

For a two-way design, the selected X-axis factor forms the chart categories. The other factor is combined with each outcome to create the bar series.

#### Outcomes on the X-axis

```text
2 = Outcomes on the X axis;
    groups/factor combinations are bar series
```

This is useful when the outcomes themselves should form the major X-axis clusters.

- In a one-way design, each group becomes a bar series.
- In a two-way design, each Factor A × Factor B combination becomes a bar series.

### 4. Two-way orientation

When the selected layout requires an X-axis factor, the macro asks which factor should form the X-axis clusters:

```text
1 = Column 1 on X axis; Column 2 as grouping series
2 = Column 2 on X axis; Column 1 as grouping series
```

This selection changes the visual orientation only. It does not alter the numerical data.

---

## Individual-value placement

Each raw observation is plotted as a circular marker over its corresponding bar.

The placement algorithm follows these rules:

1. A point is placed at the exact horizontal center of its bar by default.
2. If multiple values within the same bar are identical or visually near-identical, those points are offset horizontally.
3. The offsets are distributed around the bar center rather than shifted entirely to one side.
4. Values that do not overlap remain centered.

The current near-overlap tolerance is based on approximately **0.6% of the plotted Y-value range**. This is intended to prevent markers from completely covering one another without unnecessarily spreading all observations across the width of the bar.

These offsets are visual only. They do not change the underlying measurements, group assignments, means, SEM values, or statistical interpretation.

---

## Automatic Y-axis scaling

The macro calculates graph limits using:

- Every plotted individual value
- Each group mean minus its SEM
- Each group mean plus its SEM

It then applies approximately **8% padding** around the observed range.

Additional behavior:

- If all plotted values are nonnegative, the lower Y-axis limit is set to zero.
- If all plotted values are nonpositive, the upper Y-axis limit is set to zero.
- If all relevant values are identical, the macro creates a small nonzero range so the graph can still be displayed.
- The primary Y-axis used by the bars and the secondary Y-axis used by the individual-value scatter series receive the same minimum and maximum values.
- Secondary-axis tick labels and axis lines are hidden so the result appears as a single coordinated graph.

For separate-outcome graphs, each graph receives its own data-derived Y-axis range. Within each graph, the bars and individual observations always share the same scale.

---

## Output worksheets

### Single combined graph

The macro creates one new worksheet containing:

- A summary table of means
- A summary table of SEM values
- Sample sizes (`N`)
- A point-coordinate table used for individual-value overlays
- The completed chart

### Separate graph per outcome

The macro creates:

- One worksheet for each outcome
- One chart on each outcome worksheet
- A **Graph Index** worksheet

The Graph Index contains:

- Outcome name
- Chart worksheet name
- A hyperlink to the chart worksheet
- Source column
- PNG path, if PNG export was selected

Existing worksheets are not overwritten. If a proposed worksheet name already exists, the macro appends a number to create a unique name.

---

## PNG export

The macro can optionally export each generated chart as a PNG file.

When asked:

```text
Also export each generated chart as PNG?
```

- Choose **Yes** to select an export folder.
- Choose **No** to create the Excel charts without external image files.

Characters that are invalid in file names are replaced with underscores.

If two outcomes have identical or cleaned-to-identical names, manually verify that the exported PNG names are unique before relying on batch output.

---

## Statistical calculations

### Mean

For a group of `n` observations:

```text
Mean = sum of observations / n
```

### Sample standard deviation

The macro uses the sample standard deviation with denominator `n − 1`.

### Standard error of the mean

```text
SEM = sample standard deviation / square root of n
```

If a group contains only one valid observation:

```text
SEM = 0
```

The bars represent means, and the error bars represent **mean ± SEM**.

---

## Missing data

The macro evaluates each outcome cell independently.

- Blank cells are skipped.
- Nonnumeric outcome entries are skipped.
- A row can contribute to one outcome even if another outcome is blank.
- If a complete group/outcome combination has no valid numeric observations, the corresponding mean cell is assigned `#N/A` so Excel does not draw a misleading zero-height bar.
- Rows with blank required group or factor labels are not included.

Because nonnumeric cells are skipped silently, review the output `N` columns to confirm that the number of observations matches expectations.

---

## Troubleshooting

### The macro does not appear in the macro list

Confirm that:

- The workbook is saved as `.xlsm`.
- The `.bas` file was imported into a standard module.
- Macros are enabled.
- You are looking for `GenerateUnifiedOutcomeGraphs` in the macro list.

### Excel reports a syntax or compile error

1. Remove the older module.
2. Import the latest `UnifiedOutcomeGraphGenerator.bas` file.
3. Choose **Debug → Compile VBAProject**.
4. Make sure similarly named procedures from previous versions are not duplicated in another module.

### “No numeric observations were found”

Check that:

- The complete table, including headers, was selected.
- Outcome columns contain actual numeric values rather than numbers stored as nonnumeric text.
- Required group and factor cells are not blank.
- The correct one-way or two-way design was selected.

### The graph has too many bar series

A combined graph can become crowded when it contains many outcomes, factor levels, or factor combinations. Consider using:

```text
One SEPARATE graph per outcome
```

This usually produces a clearer figure for high-dimensional experiments.

### Raw points do not appear centered over the bars

Verify that:

- The secondary X-axis exists.
- Its minimum is `0.5`.
- Its maximum is the number of X-axis categories plus `0.5`.
- The chart has not been manually converted to a different chart type after creation.

Near-identical observations may be intentionally offset to keep overlapping points visible.

### The bars and points appear to use different vertical scales

The macro explicitly synchronizes the primary and secondary Y-axes when the chart is created. If the axes later differ, the chart may have been manually reformatted. Re-run the macro or manually assign the same minimum and maximum to both Y-axes.

### A worksheet name is different from the outcome name

Excel worksheet names are limited to 31 characters and cannot contain certain characters. The macro cleans invalid characters, truncates long names, and adds a numeric suffix when needed.

### PNG export fails

Check that:

- The selected folder is writable.
- The folder still exists.
- A file with the same name is not locked by another application.
- Your organization’s security settings permit VBA to write files to that location.

---

## Customization

The following values can be adjusted in the VBA module.

### Gap width

Search for:

```vb
ch.ChartGroups(1).GapWidth = 55
```

A larger value creates more space between category clusters. A smaller value makes bars wider.

### Marker size

Search for:

```vb
s.MarkerSize = 5
```

Increase this value for larger individual-value markers.

### Bar transparency

Search for:

```vb
s.Format.Fill.Transparency = 0.2
```

The value ranges from `0` for opaque to `1` for fully transparent.

### Y-axis padding

Search for:

```vb
pad = (yMax - yMin) * 0.08
```

Change `0.08` to adjust the amount of vertical padding.

### Near-overlap tolerance

Search for:

```vb
tol = spread * 0.006
```

Increase `0.006` to offset more near-overlapping points. Decrease it to offset only values that are extremely close.

### Color palette

Edit the RGB values in:

```vb
Private Function Palette() As Variant
```

Colors repeat if the graph contains more bar series than the palette contains colors.

---

## Important interpretation notes

- The macro produces descriptive graphs; it does not perform hypothesis testing.
- SEM describes the precision of the estimated mean and is not the same as standard deviation.
- A shared Y-axis is essential within each graph because the bars and raw points are separate Excel chart-series types.
- When outcomes have very different units or orders of magnitude, placing all outcomes on one Y-axis may make some outcomes difficult to interpret. In that situation, use separate graphs or normalize the outcomes using an analysis plan appropriate for the experiment.
- Horizontal point offsets are display adjustments only and should not be interpreted as additional variables or measured X-values.
- Always review the generated means, SEM values, and sample-size columns before using a graph in a report, presentation, or publication.

---

## Macro entry point

```vb
GenerateUnifiedOutcomeGraphs
```

---

## Version note

This README corresponds to the corrected unified module containing the expanded `UniqueSheet` function and other expanded helper procedures for improved VBA compatibility.
