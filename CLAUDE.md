# Office Add-in Project Memory (Figra)

## Project Overview

**Figra** (formerly OPEN FIGURE / FigDR) is an Excel Office Add-in that provides publication-quality scientific visualization through webR (R in WebAssembly). Users can create ggplot2-based charts directly from Excel data without installing R locally.

**Author:** Yoshiaki Sato (ORCID: 0000-0003-3375-5189)
**Platform:** Excel Office Add-in (Office.js)
**Languages:** JavaScript, R, HTML/CSS

## Repositories

**Two repositories for this project:**

1. **Figra (private)** - Main codebase
   - Local: `/Users/yoshiakisato/Desktop/github/Figra`
   - Remote: `https://github.com/h20gg702/Figra.git`
   - Contains: Add-in source code, docs (source)

2. **figra-pages (public)** - GitHub Pages for public docs
   - Local: `/Users/yoshiakisato/Desktop/github/figra-pages`
   - Remote: `https://github.com/h20gg702/figra-pages.git`
   - URL: `https://h20gg702.github.io/figra-pages/`
   - Contains: Terms of Use, Privacy Policy, Support page

**Note:** When updating docs, update both repos. Copy from `Figra/docs/` to `figra-pages/`.

## Registration System

- Registration form appears on first use (stored in localStorage as `figra_registered`)
- Submits to Google Form via hidden iframe (works in Office Add-in webview)
- Links to Terms of Use and Privacy Policy on figra-pages

**Registration Google Form:**
- URL: `https://docs.google.com/forms/d/e/1FAIpQLSczQzQZmWDl0JBtTyy0mv-5h5FsRmSyFDEIXfWXet_vs6WvKQ/formResponse`
- Entry IDs:
  - email: `entry.650578732`
  - country: `entry.838311149`
  - jobTitle: `entry.16463827`
  - affiliation: `entry.810594215`
  - agreement: `entry.438747924`

**Support Google Form:**
- URL: `https://docs.google.com/forms/d/e/1FAIpQLSeP2O069QoXC3fGLp7Js073qYzvi-iXepwbB8fvlCa554aSpw/viewform`
- Fields: Email, Issue type, OS, Excel version, Description
**Key Dependencies:**
- Office.js (Excel API)
- webR (R-wasm) - https://webr.r-wasm.org/
- ggplot2 (R package for plotting)
- svglite (R package for SVG generation)
- canvg (JavaScript library for SVG to Canvas conversion)
- vbracket (custom R package for statistical legends)

## Key Technical Concepts

### Architecture Overview

1. **Office.js Integration**: Uses Excel JavaScript API to read data and insert figures
2. **WebR Runtime**: Executes R code in browser via WebAssembly (no server required)
3. **Dual Code Generation**:
   - **Internal R Code**: Complex custom functions with 30+ parameters for maximum flexibility
   - **Educational R Code**: Simplified ggplot2 syntax for learning purposes
4. **Settings Persistence**: Stores all plot settings in `window.lastPlotSettings` when preview is clicked
5. **Tabbed UI**: 5-tab interface (Data & Chart, Colors & Style, Text & Font, Theme & Layout, Statistics)

### Custom R Functions

All internal visualizations use custom `sato_*` functions with consistent interfaces:

**Core Functions:**
- `sato_histogram()`
- `sato_box()`, `sato_box_dot()`
- `sato_violin()`, `sato_violin_dot()`
- `sato_dot()`
- `sato_bar()`, `sato_bar_error()`, `sato_bar_error_dot()`
- `sato_bar_grouped()`, `sato_bar_grouped_error()`, `sato_bar_grouped_error_dot()`
- `sato_box_grouped()`, `sato_box_grouped_dot()`
- `sato_violin_grouped()`, `sato_violin_grouped_dot()`
- `sato_line()`, `sato_line_grouped()`
- `sato_line_grouped_error()` - Pre-calculated error bars
- `sato_line_grouped_error_raw()` - Calculate statistics from raw data

**Common Parameters (shared across functions):**
- Data columns: `dat`, `group_col`, `x_col`, `y_col`, `error_col`
- Colors: `fill_color`, `stroke_color`, `fill_alpha`, `line_colors[]`
- Dimensions: `linewidth`, `barwidth`, `dodge_width`, `dot_size`, `dot_alpha`, `dot_color`, `dot_shape`
- Text: `target_font`, `title_text`, `x_text`, `y_text`, `show_title`, `show_x_label`, `show_y_label`
- Font sizes: `title_size`, `x_axis_title_size`, `y_axis_title_size`, `x_axis_text_size`, `y_axis_text_size`, `legend_text_size`
- Font weights: `title_weight`, `axis_weight`
- Scales: `x_scale`, `y_scale` (linear, log10, log2, log, -log10, -log2)
- Theme: `theme_name` (minimal, gray, bw, classic, etc.)
- Axis rotation: `x_axis_rotation`, `y_axis_rotation`, `x_axis_hjust`, `x_axis_vjust`, `y_axis_hjust`, `y_axis_vjust`
- Statistics: `add_statistics`, `statistical_test`, `variance_test`, `post_hoc_test`, `stat_symbol_type`, `stat_symbol_size`, `comparison_mode`, `custom_comparisons`, `custom_positions`
- VBracket (3+ groups): `vbracket_timepoint`, `vbracket_position`, `vbracket_x`, `vbracket_y`, `vbracket_text_size`, `vbracket_margin`, `vbracket_line_width`
- Output: `output_width`, `output_height` (for proper vbracket rendering)

### Statistical Analysis Integration

**Supported Tests:**
- **Normality Tests**: Shapiro-Wilk test (auto-selected)
- **Parametric Tests**: t-test (2 groups), ANOVA (3+ groups)
- **Non-parametric Tests**: Wilcoxon test (2 groups), Kruskal-Wallis (3+ groups)
- **Variance Tests**: Levene test, F-test
- **Post-hoc Tests**: Tukey HSD, Bonferroni, Holm, Dunnett, Dunn

**Symbol Types:**
- **Stars**: `*` (p<0.05), `**` (p<0.01), `***` (p<0.001), `ns` (not significant)
- **Letters**: Groups with different letters are significantly different (a, b, c, etc.)
- **P-values**: Shows exact p-value

**Comparison Modes:**
- **Significant only**: Shows only statistically significant comparisons
- **All comparisons**: Shows all pairwise comparisons
- **Custom selection**: User manually selects which comparisons to display with custom Y positions

### VBracket Integration (3+ Groups Line Plots)

For line plots with 3+ groups, uses custom vbracket R package to display statistical comparisons in a legend with vertical brackets.

**IMPORTANT CONCEPT: Time Point Selection**
- Statistical tests are performed at **EACH** x-axis value (e.g., Day 0, Day 3, Day 7, Day 14)
- VBracket legend shows comparisons from **ONE** selected time point only
- User must select which time point to display in the legend via dropdown
- This prevents clutter and provides clear visualization of statistical significance at a specific time

**Key Features:**
- **Time point selection**: User selects which x-axis value to show comparisons for
- Parses statistical test results filtered by selected time point
- Creates custom legend showing group colors + comparison brackets
- Supports preset positions (topleft, topright, bottomleft, bottomright)
- Supports custom X/Y positioning (0-1 scale, npc coordinates)
- Requires `output_width` and `output_height` parameters for proper rendering

**Implementation Location:**
- R function: `sato_line_grouped_error_raw()` (lines 6727-6743: filtering, lines 6745-6798: parsing)
- Filters `stat_text_results` by `vbracket_timepoint` parameter
- Parses filtered results to identify significant pairs at that time point
- Uses `legend_bracket()` from vbracket package with either `position=` or `legend_x=`/`legend_y=` parameters

**UI Controls (Statistics Tab):**
- **Visibility**: VBracket controls only appear when BOTH conditions are met:
  1. Chart type is "Grouped line with error (raw data)" (line_grouped_error_raw)
  2. "Show statistical significance" checkbox is checked
- **Hidden by default** (`display:none` in HTML line 748)
- **Dynamic visibility** controlled by `updateVbracketVisibility()` function (line 1711-1729)
- Time point selector: Dropdown populated with unique X-axis values from data (line 756-762 in HTML)
- Automatically populated when data is loaded by `populateVbracketTimepoints()` function

### Settings Storage System

**Critical Design Pattern:**
When user clicks "Preview", all UI settings are captured and stored in `window.lastPlotSettings`. This ensures:
1. R code generation reproduces the exact same figure
2. No discrepancy between internal complex code and educational simplified code
3. Consistency across preview, insert, and R code export

**Storage Location:** Line 9892-9909 in taskpane.js (inside `renderPlotPngFromSelectionFixed()`)

**Stored Settings Include:**
```javascript
window.lastPlotSettings = {
  // Chart basics
  chartType, title, xlab, ylab,
  showTitle, showXLabel, showYLabel,

  // Fonts
  fontFamily, titleSize, xAxisTitleSize, yAxisTitleSize,
  xAxisTextSize, yAxisTextSize, legendTextSize,
  titleWeight, axisWeight,

  // Colors & style
  themeName, fillColor, strokeColor, fillAlpha,
  barWidth, dodgeWidth, lineWidth,
  groupColors, dotSize, dotAlpha, dotColor, dotShape,

  // Layout
  xAxisRotation, yAxisRotation,
  xAxisHjust, xAxisVjust, yAxisHjust, yAxisVjust,
  xScale, yScale, rotation,

  // Data columns
  groupColIndex, xColIndex, yColIndex, errorColIndex,
  selectedGroupColumn, selectedXColumn, selectedYColumn, selectedErrorColumn,

  // Statistics
  addStatistics, errorBarType, statisticalTest, varianceTest, postHocTest,
  statSymbolType, statSymbolSize, comparisonMode,
  customComparisons, customPositions,

  // VBracket (3+ groups line plots)
  vbracketTimepoint, vbracketPosition, vbracketX, vbracketY,
  vbracketTextSize, vbracketMargin, vbracketLineWidth,

  // Data ordering
  dataOrder, customOrderGroup, customOrderCategory,

  // Other
  numGroups, bins
};
```

### Data Ordering System

**Three ordering modes:**
1. **Original order**: Preserves Excel data order exactly as entered
2. **Alphabetical**: Default ggplot2 alphabetical sorting
3. **Custom order**: User can drag-and-drop to reorder both groups and categories

**Implementation:**
- Groups: Reordered for grouped charts (e.g., bar_grouped, line_grouped)
- Categories: Reordered for X-axis values
- Stored in: `customOrderGroup` and `customOrderCategory` arrays

### Text Formatting System

Supports rich text formatting in titles and axis labels:
- **Italic**: `*italic text*` → rendered as italic
- **Superscript**: `^superscript^` → rendered as superscript
- **Subscript**: `~subscript~` → rendered as subscript
- **Preset symbols**: log₂, log₁₀ buttons for quick insertion

**Implementation:**
- UI buttons call `formatText()` and `insertSymbol()` JavaScript functions
- Converted to R plotmath expressions in R code generation

## File Structure

### HTML: `/Users/yoshiakisato/My Office Add-in/src/taskpane/taskpane.html`

**Purpose:** UI layout with tabbed interface

**Key Sections:**

**Lines 1-129: Head & Styles**
- Office.js CDN import (line 11)
- CSS for tabs, buttons, inputs, formatting
- Color picker styles (lines 60-82)
- Tab navigation styles (lines 84-127)

**Lines 130-142: Title & Tab Navigation**
- App title "OPEN FIGURE"
- 5 tab buttons: Data & Chart, Colors & Style, Text & Font, Theme & Layout, Statistics

**Lines 144-278: Tab 1 - Data & Chart**
- Chart type selector (lines 147-171): 18 chart types
- Bins input for histogram (lines 174-177)
- Group/X/Y/Error column selectors (lines 180-225)
- Data ordering controls with drag-and-drop (lines 196-219)
- Action buttons (lines 227-245):
  - Load Data, Preview, Insert Figure to Sheet
  - Download R Code, R Code to Cell
  - Export statistics to selected cells (line 238-242)
- Export settings: width, height, units, DPI (lines 248-259)
- Preset save/load system (lines 262-276)

**Lines 280-407: Tab 2 - Colors & Style**
- Bar interior: fill color and opacity (lines 284-297)
- Bar border: stroke color (lines 300-307)
- Group colors: Dynamic controls for 2-6 groups (lines 310-359)
- Line width, bar width, dodge width (lines 362-377)
- Dot settings: size, color, shape, alpha (lines 380-405)

**Lines 409-525: Tab 3 - Text & Font**
- Font family selector (lines 413-424)
- Title section: show/hide, text area with format buttons, size, weight (lines 426-454)
- X axis section: show/hide, text area with format buttons, title size, text size (lines 456-481)
- Y axis section: show/hide, text area with format buttons, title size, text size (lines 483-508)
- Legend text size (lines 511-514)
- Axis label weight (lines 517-523)
- Format buttons: italic, superscript, subscript, log₂, log₁₀

**Lines 527-622: Tab 4 - Theme & Layout**
- ggtheme selector: gray, bw, minimal, classic, etc. (lines 531-544)
- X/Y axis scales: linear, log10, log2, ln, -log10, -log2 (lines 546-570)
- Axis line width (lines 573-576)
- X/Y range inputs (lines 579-592)
- Chart rotation: 0° or 90° (lines 595-601)
- Axis text positioning: rotation, hjust, vjust for both axes (lines 604-620)

**Lines 624-830: Tab 5 - Statistics**

**Lines 628-645: Unsupported Chart Type Message**
- Warning message shown when selected chart type doesn't support statistics
- Lists all 8 supported chart types:
  - Box plot, Box plot + dots, Violin plot + dots, Bar + error + dots
  - Grouped bar + dot plot with error, Grouped box plot with dots
  - Grouped violin plot with dots, Grouped line with error (raw data)
- Hidden by default (`display:none`)
- Controlled by `updateStatisticsAvailability()` JS function

**Lines 648-833: Statistics Controls Container** (`statsControlsContainer`)
- Wraps ALL statistics controls
- Only shown when chart type supports statistics
- Controlled by `updateStatisticsAvailability()` JS function

- **Statistical significance checkbox (lines 649-652)** - Always visible when chart type supports stats

**Lines 655-831: Detailed Statistics Controls** (`statsDetailedControls`)
- Wraps all detailed configuration controls
- Hidden by default (`display:none`)
- Only shown when "Show statistical significance" checkbox is checked
- Controlled by `updateStatsDetailedControlsVisibility()` JS function
- **Progressive disclosure**: Reduces UI clutter by showing controls only when needed

**Detailed Controls Include:**
- Error bar type: SD, SE, CI95
- Display comparisons mode: significant only, all, custom
- Statistical test selector
- Variance test: Levene, F-test
- Symbol type: stars, letters, p-value
- Significance levels display
- Post-hoc test selector
- Dunnett control group selector
- **Traditional stat controls** - wrapped in `traditionalStatControls` div:
  - **Display options**:
    - Show overall test symbol (ANOVA/Kruskal-Wallis) checkbox
    - Show pairwise comparison lines checkbox
    - Pairwise selection controls
  - **Visual controls**:
    - Symbol size, line thickness, distance from bars, vertical position adjustment
  - Reset to defaults button
  - **Hidden when VBracket controls are shown** (for 3+ groups line plots)
  - **Visible for all other chart types** with statistics enabled
- **VBracket Legend Controls**:
  - **Visibility**: Hidden by default (`display:none`), only shown when:
    - Chart type = "line_grouped_error_raw" AND
    - "Show statistical significance" is checked
    - Controlled by `updateVbracketVisibility()` JS function
  - **Time point selector (lines 779-785)**: Dropdown populated with unique X-axis values
    - Automatically filled when data is loaded
    - Only applicable for line_grouped_error_raw chart type
    - Selects which time point's comparisons to show in legend
  - **Position selector (lines 788-796)**: topleft, topright, bottomleft, bottomright, custom
  - **X position (lines 798-802)**: 0-1 scale (0=left, 1=right)
    - Hidden by default, only shown when position = "custom"
    - Controlled by `updateVbracketPositionInputs()` JS function (line 1771-1782)
  - **Y position (lines 804-808)**: 0-1 scale (0=bottom, 1=top)
    - Hidden by default, only shown when position = "custom"
    - Controlled by `updateVbracketPositionInputs()` JS function (line 1771-1782)
  - Text size: 6-20
  - Margin: 0.05-0.5 (distance from plot edge)
  - Line width: 0.1-2

**Lines 832-878: Status, Canvas, Scripts**
- Status message area (line 834)
- Debug panel (lines 836-839)
- Canvas for preview (line 841)
- Tab switching JavaScript (lines 845-865)
- WebR module import (lines 869-872)
- taskpane.js import (line 875)

### JavaScript: `/Users/yoshiakisato/My Office Add-in/src/taskpane/taskpane.js`

**Purpose:** Main application logic, R code generation, Excel integration

**File Size:** 473KB (10,000+ lines)

**Key Sections:**

**Lines 1-100: Global Variables & Initialization**
- `window.webR` - WebR instance
- `window.lastProcessedData` - Stores loaded Excel data
- `window.lastPlotSettings` - Stores all settings from last preview
- Office.onReady() initialization

**Lines 100-500: UI Event Handlers**
- Load Data button
- Preview button
- Insert Figure button
- Download R Code button
- R Code to Cell button
- Export Statistics button
- Tab switching
- Color picker synchronization
- Format text buttons

**Lines 1712-1721: STATS_SUPPORTED_CHART_TYPES Constant**
Array of chart types that support statistical analysis:
- `box`, `box_dot`, `violin_dot`, `bar_error_dot`
- `bar_grouped_error_dot`, `box_grouped_dot`, `violin_grouped_dot`
- `line_grouped_error_raw`

**Lines 1724-1742: updateStatisticsAvailability() Function**
Controls visibility of entire Statistics tab content based on chart type support:
- Shows statistics controls container when chart type is in STATS_SUPPORTED_CHART_TYPES
- Shows warning message with list of supported types when chart type doesn't support statistics
- Called by:
  - Chart type change event (line 1872)
  - Initial page load via handleChartTypeChange()

**Lines 1745-1753: updateStatsDetailedControlsVisibility() Function**
Controls visibility of detailed statistics controls based on checkbox state:
- Shows all detailed controls when "Show statistical significance" checkbox is checked
- Hides all detailed controls when checkbox is unchecked
- **Progressive disclosure** - reduces UI clutter by showing controls only when needed
- Called by:
  - Statistics checkbox change event (line 1886)
  - Chart type change (line 1875) to ensure correct state on chart type switch
  - Initial page load via handleChartTypeChange()

**Lines 1755-1779: updateVbracketVisibility() Function**
Controls visibility of VBracket controls section AND traditional stat controls:
- Shows VBracket controls only when chart type = "line_grouped_error_raw" AND statistics checkbox is checked
- **When VBracket is shown**: Hides traditional stat controls (symbol size, line thickness, distance from bars, vertical adjustment)
- **When VBracket is hidden**: Shows traditional stat controls
- Automatically calls `populateVbracketTimepoints()` when showing vbracket controls
- Called by:
  - Chart type change event (line 1878)
  - Statistics checkbox change event (line 1887)

**Lines 1781-1792: updateVbracketPositionInputs() Function**
Controls visibility of X/Y position inputs:
- Shows X/Y position inputs only when vbracket position dropdown = "custom"
- Hides them for preset positions (topleft, topright, etc.)
- Called by:
  - vbracketPosition dropdown change event (line 1891)

**Lines 1809-1812: Error Column Support**
Define which chart types need error column selector (pre-calculated error data):
- `line_grouped_error` - Group, X, Y, Error (4 columns)
- `bar_error` - X, Y, Error (3 columns)
- `bar_grouped_error` - Group, X, Y, Error (4 columns)

**Lines 10854-10859: Auto-detect Error Column**
Automatically assigns error column when data is loaded:
- For grouped error charts with 4+ columns: Error = 4th column
- For single-group bar error with 3+ columns: Error = 3rd column

**Lines 500-1500: Data Loading & Processing**
- Excel data reading via Office.js
- Data validation and cleaning
- Column detection and mapping
- Group/category extraction
- Custom ordering implementation

**Lines 1282-1288: collectCurrentSettings() - VBracket Settings**
Collects current UI values for vbracket parameters:
```javascript
vbracketPosition: el("vbracketPosition")?.value || "topright",
vbracketX: el("vbracketX")?.value || "0.85",
vbracketY: el("vbracketY")?.value || "0.85",
vbracketTextSize: el("vbracketTextSize")?.value || "10",
vbracketMargin: el("vbracketMargin")?.value || "0.15",
vbracketLineWidth: el("vbracketLineWidth")?.value || "0.5",
```

**Lines 808-927: extractRelevantRCode() - Educational R Code Generation**
**CRITICAL FUNCTION:** Generates simplified R code that reproduces the preview figure.

**Key Implementation:**
1. Checks for `window.lastPlotSettings` - if missing, prompts user to click Preview first
2. Uses stored settings instead of re-reading UI (prevents mismatch)
3. Reconstructs data frame from `window.lastProcessedData`
4. Calls `generateEducationalRCode()` with stored settings
5. Returns executable R code string

**Purpose:** Allows users to learn basic ggplot2 syntax while ensuring code reproduces exact figure

**Lines 2000-8000: R Function Definitions (Custom sato_* Functions)**

All R functions follow similar structure:
```r
sato_CHARTTYPE <- function(dat, group_col, x_col, y_col, ...) {
  # 1. Data preparation
  # 2. Statistical analysis (if add_statistics=TRUE)
  # 3. ggplot2 plot construction
  # 4. Theme application
  # 5. Statistical annotations (if requested)
  # 6. VBracket legend (for 3+ groups line plots)
  # 7. Return plot object
}
```

**Lines 6103-6125: sato_line_grouped_error_raw() Signature**
Function for grouped line plots with error bars calculated from raw data:
```r
sato_line_grouped_error_raw <- function(
  dat, group_col=1, x_col=2, y_col=3,
  line_colors=c("#4C78A8", "#E15759", "#76B7B2", "#F28E2B"),
  linewidth=1.5, alpha=0.9, ribbon_alpha=0.2,
  error_type="sd", use_ribbon=TRUE,
  group_name="Group", x_name="X", value_name="Value",
  target_font="Arial", title_weight="plain", axis_weight="plain",
  title_size=14, x_axis_title_size=12, y_axis_title_size=12,
  x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
  title_text="Grouped Line Plot with Error", x_text="X", y_text="Value",
  show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
  x_scale="linear", y_scale="linear",
  theme_name="minimal",
  x_axis_rotation=0, y_axis_rotation=0,
  x_axis_hjust=0.5, x_axis_vjust=0.5,
  y_axis_hjust=0.5, y_axis_vjust=0.5,
  add_statistics=FALSE, statistical_test="auto",
  variance_test="levene", stat_symbol_size=7,
  comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
  vbracket_position="topright", vbracket_x=0.85, vbracket_y=0.85,
  vbracket_text_size=10, vbracket_margin=0.15, vbracket_line_width=0.5,
  output_width=6, output_height=4
)
```

**Lines 6646-6738: VBracket Implementation for 3+ Groups**
Critical implementation for showing statistical comparisons in custom legend:

```r
} else if (n_groups >= 3) {
  cat("Multi-group line plot (3+ groups): Using vbracket for statistical legend\\n")

  if (!requireNamespace("vbracket", quietly = TRUE)) {
    cat("⚠️ vbracket package not installed.\\n")
  } else {
    tryCatch({
      library(vbracket)

      # Parse stat_text_results to extract significant comparisons
      significant_comparisons <- list()
      comparison_labels <- c()

      lines <- strsplit(stat_text_results, "\\n")[[1]]
      for (line in lines) {
        # Look for lines like "Group1 vs Group2: * (p=0.023)"
        if (grepl("vs", line) && grepl("\\*", line)) {
          parts <- strsplit(line, ":")[[1]]
          if (length(parts) >= 2) {
            groups_part <- trimws(parts[1])
            groups <- strsplit(groups_part, " vs ")[[1]]
            if (length(groups) == 2) {
              group1 <- trimws(groups[1])
              group2 <- trimws(groups[2])

              # Extract symbol (*, **, ***)
              symbol_match <- regmatches(parts[2], regexpr("\\*+", parts[2]))
              symbol <- if(length(symbol_match) > 0) symbol_match else "*"

              significant_comparisons[[length(significant_comparisons) + 1]] <-
                c(group1, group2)
              comparison_labels <- c(comparison_labels, symbol)
            }
          }
        }
      }

      if (length(significant_comparisons) > 0) {
        group_labels <- levels(plot_data$group)
        legend_colors <- line_colors[1:length(group_labels)]

        p <- p + theme(legend.position = "none")

        # Use custom position if specified
        if (vbracket_position == "custom") {
          p <- p + legend_bracket(
            labels = group_labels,
            colors = legend_colors,
            comparisons = significant_comparisons,
            comparison_labels = comparison_labels,
            legend_x = vbracket_x,
            legend_y = vbracket_y,
            text_size = vbracket_text_size,
            margin = vbracket_margin,
            line_width = vbracket_line_width,
            output_width = output_width,
            output_height = output_height,
            text_family = target_font
          )
        } else {
          p <- p + legend_bracket(
            labels = group_labels,
            colors = legend_colors,
            comparisons = significant_comparisons,
            comparison_labels = comparison_labels,
            position = vbracket_position,
            text_size = vbracket_text_size,
            margin = vbracket_margin,
            line_width = vbracket_line_width,
            output_width = output_width,
            output_height = output_height,
            text_family = target_font
          )
        }
      }
    }, error = function(e) {
      cat("Error adding vbracket:", e$message, "\\n")
    })
  }
}
```

**Lines 8000-9000: Statistical Analysis Functions**
- Normality testing (Shapiro-Wilk)
- T-test and Wilcoxon test implementation
- ANOVA and Kruskal-Wallis implementation
- Post-hoc test implementations (Tukey HSD, Bonferroni, etc.)
- Result parsing and formatting
- Bracket positioning calculations

**Lines 9000-9800: Plot Rendering & Excel Integration**
- SVG generation via svglite
- Canvas conversion via canvg
- PNG generation from canvas
- Excel shape insertion
- Error handling and status updates

**Lines 9805-9812: R Function Call with VBracket Parameters**
Example of passing vbracket settings to R function:
```javascript
vbracket_position = "${vbracketPosition}",
vbracket_x = ${vbracketX},
vbracket_y = ${vbracketY},
vbracket_text_size = ${vbracketTextSize},
vbracket_margin = ${vbracketMargin},
vbracket_line_width = ${vbracketLineWidth},
output_width = ${wIn},
output_height = ${hIn}
```

**Lines 9892-9909: Settings Storage in renderPlotPngFromSelectionFixed()**
**CRITICAL:** Stores all settings when Preview is clicked:
```javascript
// Store all settings for R code generation
window.lastPlotSettings = {
  chartType, title, xlab, ylab,
  showTitle: showTitle === 'TRUE',
  showXLabel: showXLabel === 'TRUE',
  showYLabel: showYLabel === 'TRUE',
  fontFamily: rFontName, titleSize, xAxisTitleSize, yAxisTitleSize,
  xAxisTextSize, yAxisTextSize, legendTextSize, titleWeight, axisWeight,
  themeName, fillColor, strokeColor, fillAlpha, barWidth, dodgeWidth, lineWidth,
  groupColors, dotSize, dotAlpha, dotColor, dotShape,
  xAxisRotation, yAxisRotation, xAxisHjust, xAxisVjust, yAxisHjust, yAxisVjust,
  xScale, yScale, rotation, groupColIndex, xColIndex, yColIndex, errorColIndex,
  selectedGroupColumn, selectedXColumn, selectedYColumn, selectedErrorColumn,
  addStatistics, errorBarType, statisticalTest, varianceTest, postHocTest,
  statSymbolType, statSymbolSize, comparisonMode, customComparisons, customPositions,
  vbracketPosition, vbracketX, vbracketY, vbracketTextSize, vbracketMargin, vbracketLineWidth,
  dataOrder, customOrderGroup, customOrderCategory, numGroups, bins
};
console.log('📦 Stored plot settings for R code generation:', window.lastPlotSettings);
```

**Lines 10000+: Utility Functions**
- Color validation and conversion
- Font family mapping (JavaScript ↔ R)
- Column name sanitization
- Preset save/load implementation
- Export statistics to Excel cells

## Important Design Decisions

### 1. Settings Storage Strategy (CRITICAL)

**Problem:** The add-in uses complex internal R code (30+ parameters) but needs to export simple educational R code. Initial approach re-read UI values when generating R code, causing mismatches.

**Solution:** Store ALL settings in `window.lastPlotSettings` when Preview is clicked. R code generation uses these stored values, ensuring perfect reproduction.

**Implementation:** Lines 9892-9909 (storage), Lines 808-927 (usage)

**Impact:** Guarantees that "R code to cell" produces identical figure to preview.

### 2. VBracket Integration (3+ Groups Only)

**Decision:** Use vbracket custom legend only for line plots with 3+ groups where standard statistical brackets become cluttered.

**Rationale:**
- 2 groups: Simple asterisks above bars work well
- 3+ groups: Multiple comparisons need visual organization

**Implementation:** Lines 6646-6738 in `sato_line_grouped_error_raw()`

**Requirements:**
- Parses statistical test results text to extract significant pairs
- Must pass `output_width` and `output_height` for proper rendering
- Supports both preset positions and custom X/Y coordinates

### 3. Custom X/Y Positioning for VBracket

**Decision:** Added "Custom (X/Y)" option alongside preset positions (topleft, topright, etc.)

**Rationale:** Preset positions sometimes don't account for plot-specific layouts (e.g., long axis labels, specific data distributions)

**Implementation:**
- UI controls: Lines 755-776 in taskpane.html
- R code: Uses `legend_x`/`legend_y` when `vbracket_position == "custom"`
- Coordinate system: 0-1 scale (0,0 = bottom-left, 1,1 = top-right)

### 4. No Export of print.gg() Method

**Decision:** Do not export custom `print.gg()` method in generated R code.

**Rationale:**
- Dangerous to override base R methods in educational code
- vbracket package provides `ggsave_vbracket()` for METHOD 2
- Add-in uses METHOD 1 (annotation_custom with makeContent) which works with regular `ggsave()`

### 5. Tab-Based UI Organization

**Decision:** Split all controls into 5 thematic tabs instead of one long scrolling page.

**Rationale:**
- Reduces cognitive load
- Logical grouping of related controls
- Better mobile/small-screen support
- Statistics tab keeps advanced features separate

**Implementation:** Lines 134-141 (tab navigation), separate div sections for each tab

### 6. Data Ordering System

**Decision:** Three modes (Original, Alphabetical, Custom) instead of forcing alphabetical or custom only.

**Rationale:**
- Scientists often enter data in meaningful order (e.g., time sequence, treatment escalation)
- Original order preserves user intent
- Custom order allows fine-tuning for publication
- Alphabetical needed for consistency with ggplot2 defaults

**Implementation:** Lines 196-219 in HTML, custom drag-and-drop in JavaScript

### 7. Font Family Mapping

**Decision:** Map JavaScript font names to R font names that work cross-platform.

**Rationale:**
- Windows/Mac have different font availability
- R (via svglite) needs specific font family names
- Fallback to safe defaults prevents rendering errors

**Mapping Example:**
```javascript
"Arial" → "Arial"
"Times New Roman" → "Times"
"Noto Sans JP" → "Noto Sans CJK JP"
```

## Chart Type Reference

### Single-Group Charts

**histogram**
- Function: `sato_histogram()`
- Requires: Y column (values)
- Optional: bins parameter
- Statistics: N/A

**box** / **box_dot**
- Function: `sato_box()` / `sato_box_dot()`
- Requires: X column (categories), Y column (values)
- Statistics: Comparisons between categories

**violin** / **violin_dot**
- Function: `sato_violin()` / `sato_violin_dot()`
- Requires: X column (categories), Y column (values)
- Statistics: Comparisons between categories

**dot**
- Function: `sato_dot()`
- Requires: X column (categories), Y column (values)
- Statistics: Comparisons between categories

**bar** / **bar_error** / **bar_error_dot**
- Function: `sato_bar()` / `sato_bar_error()` / `sato_bar_error_dot()`
- Requires: X column (categories), Y column (values)
- Optional: Error column (for pre-calculated errors)
- Statistics: Comparisons between categories

**line**
- Function: `sato_line()`
- Requires: X column (continuous/categorical), Y column (values)
- Statistics: N/A (single series)

### Grouped Charts

**bar_grouped** / **bar_grouped_error** / **bar_grouped_error_dot**
- Function: `sato_bar_grouped()` / `sato_bar_grouped_error()` / `sato_bar_grouped_error_dot()`
- Requires: Group column, X column (categories), Y column (values)
- Optional: Error column (for pre-calculated errors)
- Statistics: Pairwise comparisons between groups within each category

**box_grouped** / **box_grouped_dot**
- Function: `sato_box_grouped()` / `sato_box_grouped_dot()`
- Requires: Group column, X column (categories), Y column (values)
- Statistics: Pairwise comparisons between groups within each category

**violin_grouped** / **violin_grouped_dot**
- Function: `sato_violin_grouped()` / `sato_violin_grouped_dot()`
- Requires: Group column, X column (categories), Y column (values)
- Statistics: Pairwise comparisons between groups within each category

**line_grouped**
- Function: `sato_line_grouped()`
- Requires: Group column, X column (continuous/categorical), Y column (values)
- Statistics: N/A (shows trends only)

**line_grouped_error**
- Function: `sato_line_grouped_error()`
- Requires: Group column, X column, Y column (mean values), Error column (4 columns total)
- Statistics: Comparisons at each X point (2 groups) or overall (3+ groups)
- Note: Pre-calculated mean and error values required

**bar_error**
- Function: `sato_bar_error()`
- Requires: X column (categories), Y column (mean values), Error column (3 columns total)
- Statistics: Not supported (no raw data points)
- Note: Pre-calculated mean and error values required

**bar_grouped_error**
- Function: `sato_bar_grouped_error()`
- Requires: Group column, X column (categories), Y column (mean values), Error column (4 columns total)
- Statistics: Not supported (no raw data points)
- Note: Pre-calculated mean and error values required

**line_grouped_error_raw**
- Function: `sato_line_grouped_error_raw()`
- Requires: Group column, X column, Y column (raw values)
- Calculates: Mean and error from raw data automatically
- Statistics: Full statistical analysis with post-hoc tests
- **VBracket Support**: For 3+ groups, uses vbracket for custom legend with comparison brackets
- Note: Most powerful option for line plots with statistical analysis

### Specialized Charts

**ic50_dose_response**
- Function: 4-Parameter Logistic (4PL) model fitting
- Requires: X column (concentration), Y column (response)
- Model: `response = Bottom + (Top - Bottom) / (1 + (conc/IC50)^Hill)`
- Features:
  - Automatic IC50 calculation
  - Curve style customization (color, width)
  - IC50 reference lines (horizontal at 50%, vertical at IC50)
  - Data display options: all points, mean±SD, mean±SE
  - Export IC50 results to new sheet
- UI Controls:
  - IC50 Curve Style section (Colors & Style tab)
  - IC50 Analysis section (Statistics tab)
- Fallback: If 4PL fails, attempts 3PL (fixed Bottom = 0)

## Typical Usage Patterns

### Basic Bar Plot with Statistics

```javascript
// 1. User loads data from Excel (3 columns: Group, Category, Value)
// 2. Select chart type: "Bar + error + dots"
// 3. Map columns: X = Category, Y = Value
// 4. Go to Statistics tab
// 5. Check "Show statistical significance"
// 6. Select test: "Auto-select optimal test"
// 7. Select symbol type: "Stars"
// 8. Click Preview
// 9. Click "Insert Figure to Sheet"
```

### Grouped Line Plot with VBracket Legend (3+ Groups)

```javascript
// 1. User loads data (4 columns: Group, Time, Value, Replicate)
//    Example groups: Control, Treatment A, Treatment B
// 2. Select chart type: "Grouped line with error (raw data)"
// 3. Map columns: Group = Group, X = Time, Y = Value
// 4. Go to Statistics tab
// 5. Check "Show statistical significance"
// 6. Select comparison mode: "Show significant only"
// 7. Configure VBracket legend:
//    - Position: "Top Right" or "Custom (X/Y)"
//    - If custom: X = 0.85, Y = 0.85
//    - Text size: 10
//    - Margin: 0.15
//    - Line width: 0.5
// 8. Click Preview → vbracket legend appears showing significant comparisons
// 9. Adjust position if needed
// 10. Click "Insert Figure to Sheet"
```

### Custom Comparison Selection

```javascript
// 1. Load grouped data
// 2. Go to Statistics tab
// 3. Check "Show statistical significance"
// 4. Select "Custom selection" comparison mode
// 5. Check specific comparisons you want to show
// 6. Adjust Y position sliders for each comparison to avoid overlap
// 7. Click Preview
// 8. Iterate adjustments as needed
```

### Exporting R Code for Reproduction

```javascript
// 1. Create your plot and click Preview
// 2. Verify the preview looks correct
// 3. Click "R Code to Cell"
// 4. R code is inserted to selected Excel cell
// 5. Copy code to R/RStudio
// 6. Run code → produces identical figure
```

## Common Issues & Solutions

### Issue 1: R Code Doesn't Reproduce Figure

**Symptom:** "R code to cell" output creates different figure than preview

**Root Cause:** User changed UI settings after clicking Preview but before generating R code

**Solution:**
1. Always click Preview first
2. Don't change any settings after preview
3. If settings changed, click Preview again before generating R code

**Technical Explanation:** R code uses `window.lastPlotSettings` which is only updated when Preview is clicked (line 9892-9909)

### Issue 2: VBracket Legend Not Appearing

**Symptom:** Statistical analysis runs but no vbracket legend appears in 3+ group line plot

**Root Cause:** One of several possibilities:
1. No significant comparisons found
2. vbracket package not installed in webR
3. Chart type is not "Grouped line with error (raw data)"
4. Statistics not enabled

**Solution:**
1. Check R console output for vbracket installation status
2. Verify chart type is `line_grouped_error_raw`
3. Ensure "Show statistical significance" is checked
4. Try "Show all comparisons" mode to see if only non-significant results exist

### Issue 3: VBracket Legend Overlaps Plot Data

**Symptom:** Legend obscures data points or lines

**Solution:**
1. Change position from preset (e.g., "Top Right") to "Custom (X/Y)"
2. Adjust X and Y position sliders (0-1 scale)
3. Increase margin parameter to move legend further from edge
4. Decrease text size if needed

**Optimal Settings (general guidance):**
- Top right: X = 0.85-0.95, Y = 0.85-0.95
- Top left: X = 0.05-0.15, Y = 0.85-0.95
- Bottom right: X = 0.85-0.95, Y = 0.05-0.15
- Bottom left: X = 0.05-0.15, Y = 0.05-0.15

### Issue 4: Statistical Brackets Overlap

**Symptom:** Multiple comparison brackets overlap each other

**Solution:**
1. Use "Custom selection" comparison mode
2. Show only most important comparisons
3. Adjust Y position for each comparison using sliders
4. For 3+ groups line plots, use vbracket legend instead (automatically handles spacing)

### Issue 5: Browser Cache Showing Old UI

**Symptom:** UI changes not appearing, seeing old version of controls

**Root Cause:** Office Add-in cached old HTML/JS files

**Solution:**
1. Close Excel completely
2. Reopen Excel and load add-in
3. If still not updated, try Ctrl+F5 (hard refresh) in add-in pane
4. Clear Office add-in cache (platform-specific)

### Issue 6: Font Not Rendering Correctly

**Symptom:** Text appears in wrong font or falls back to default

**Root Cause:** Font not available in R/svglite or OS

**Solution:**
1. Use cross-platform fonts: Arial, Times, Courier
2. For Japanese: Use "Noto Sans JP" (included in webR)
3. Avoid platform-specific fonts (e.g., "SF Pro" on Mac)

### Issue 7: Figure Size Mismatch

**Symptom:** Figure appears stretched or compressed when inserted

**Root Cause:** Export settings (width, height, DPI) don't match intended display size

**Solution:**
1. Set export dimensions in Data & Chart tab (default: 6×4 inches, 300 DPI)
2. For vbracket plots, ensure these match the `output_width` and `output_height` parameters
3. Standard sizes:
   - Presentation: 10×7.5 inches, 150 DPI
   - Publication: 6×4 inches or 7×5 inches, 300-600 DPI
   - Poster: 12×9 inches, 300 DPI

## Development Workflow

### Local Development

1. Edit `taskpane.html` or `taskpane.js`
2. Save changes
3. In Excel: Close add-in pane
4. Reopen add-in (Forces reload)
5. Test changes
6. Repeat

**Note:** May need to clear browser cache for major changes

### Testing Chart Types

Use test data with known properties:
- **Normal distribution:** For testing parametric tests
- **Skewed distribution:** For testing non-parametric tests
- **Equal variance:** For testing standard t-tests
- **Unequal variance:** For testing Welch's t-test
- **Multiple groups:** For testing ANOVA and post-hoc tests

### Testing VBracket Integration

1. Create 3-6 groups of data with known significant differences
2. Select `line_grouped_error_raw` chart type
3. Enable statistics
4. Test all position options (presets + custom)
5. Verify legend appears correctly
6. Verify brackets connect correct groups
7. Verify symbols match statistical results
8. Export R code and verify reproduction

### Debugging Tips

**Enable Debug Panel:**
```javascript
document.getElementById('debugPanel').style.display = 'block';
document.getElementById('debugContent').textContent = 'Debug info here';
```

**Check Console Logs:**
- Look for "📦 Stored plot settings..." message after Preview
- Check for R execution errors
- Verify webR initialization

**Common Console Messages:**
```
✓ webR initialized successfully
✓ Loaded 1000 rows of data
📦 Stored plot settings for R code generation
⚠️ vbracket package not installed
❌ Error in statistical test: ...
```

## Future Development Notes

### Planned Features (mentioned in code but not implemented)

1. **Scatter plots**: Chart type option exists but function not implemented
2. **4PL curves**: Mentioned in comments, for dose-response curves
3. **More themes**: Could add ggthemes package themes
4. **Export to PowerPoint**: Direct PPT export instead of copy-paste
5. **Preset sharing**: Import/export preset JSON files

### VBracket Enhancements

1. **Auto-position optimization**: Automatically find best position to avoid overlap
2. **Multiple legends**: Support for comparing different facets
3. **Horizontal brackets**: For certain plot orientations
4. **Symbol customization**: Allow user to override auto-generated symbols

### Statistical Analysis Enhancements

1. **Effect size reporting**: Add Cohen's d, eta-squared, etc.
2. **Power analysis**: Sample size recommendations
3. **Mixed models**: For repeated measures
4. **Survival analysis**: Kaplan-Meier curves with log-rank test

### Code Generation Improvements

1. **Generate tidyverse code**: Option for dplyr/tidyr preprocessing
2. **Generate publication script**: Complete reproducible analysis
3. **Generate method section**: Auto-write statistical methods text
4. **Export data subset**: Save processed data with code

## External Dependencies

### R Packages (webR)

**Core (always loaded):**
- ggplot2 >= 3.0.0
- svglite
- grid (base R)

**Statistical (loaded on demand):**
- car (Levene test)
- stats (base R, t-test, ANOVA)
- PMCMRplus (Dunn test)
- multcomp (Dunnett test)

**Custom:**
- vbracket (custom package, user-developed)
  - Repository: https://github.com/yoshiakisato/vbracket
  - Installation in webR: Not automatic, requires manual setup
  - Location: /Users/yoshiakisato/Desktop/github/vbracket_github

### JavaScript Libraries

**Office.js:**
- CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js
- Version: Latest (dynamic)
- Purpose: Excel API access

**webR:**
- CDN: https://webr.r-wasm.org/latest/webr.mjs
- Version: Latest
- Purpose: R runtime in browser

**canvg:**
- Loaded dynamically via renderPlotPngFromSelectionFixed()
- Purpose: SVG to Canvas conversion for PNG generation

## Quick Reference: Key Variable Names

### Data Storage
- `window.webR` - WebR instance
- `window.lastProcessedData` - Array of data rows from Excel
- `window.lastPlotSettings` - Complete settings object from last Preview

### HTML Element IDs (Commonly Referenced)

**Chart & Data:**
- `chartType` - Chart type dropdown
- `groupColumn` - Group column selector
- `xColumn` - X axis column selector
- `yColumn` - Y axis column selector
- `errorColumn` - Error column selector
- `dataOrder` - Data ordering mode
- `numGroups` - Number of groups for color mapping

**Colors:**
- `fillColor` / `fillPicker` - Bar fill color
- `strokeColor` / `strokePicker` - Bar border color
- `fillAlpha` - Fill opacity
- `groupColor1`-`groupColor6` - Individual group colors
- `dotColor` / `dotColorPicker` - Dot color
- `dotSize`, `dotAlpha`, `dotShape` - Dot appearance

**Text:**
- `fontFam` - Font family
- `titleText` / `xLabel` / `yLabel` - Axis labels
- `showTitle` / `showXLabel` / `showYLabel` - Show/hide checkboxes
- `titleSize`, `xAxisTitleSize`, `yAxisTitleSize` - Title sizes
- `xAxisTextSize`, `yAxisTextSize` - Axis text sizes
- `legendTextSize` - Legend size
- `titleWeight` / `axisWeight` - Bold/plain

**Theme:**
- `ggtheme` - ggplot2 theme
- `xScale` / `yScale` - Axis scales
- `rotation` - Chart rotation (0° or 90°)
- `xAxisRotation`, `yAxisRotation` - Axis text rotation
- `xAxisHjust`, `xAxisVjust`, `yAxisHjust`, `yAxisVjust` - Text alignment

**Statistics:**
- `addStatistics` - Enable statistics checkbox
- `errorBarType` - SD, SE, or CI95
- `statisticalTest` - Test selection
- `varianceTest` - Levene or F-test
- `postHocTest` - Post-hoc test for 3+ groups
- `statSymbolType` - Stars, letters, or p-value
- `statSymbolSize` - Symbol size
- `comparisonMode` - significant/all/custom
- `showPairwiseComparisons` - Show bracket lines

**VBracket (3+ groups line plots):**
- `vbracketPosition` - Preset position or "custom"
- `vbracketX` - X position (0-1 scale)
- `vbracketY` - Y position (0-1 scale)
- `vbracketTextSize` - Legend text size
- `vbracketMargin` - Distance from plot edge
- `vbracketLineWidth` - Bracket line width

**Export:**
- `expWidth` / `expHeight` - Figure dimensions
- `expUnits` - in, cm, or mm
- `expDpi` - Resolution

**Actions:**
- `load` - Load Data button
- `preview` - Preview button
- `insert` - Insert Figure button
- `downloadRCode` - Download R Code button
- `writeRCodeToCell` - R Code to Cell button
- `exportStatResults` - Export statistics button
- `savePreset` / `loadPreset` - Preset management

## Related Projects

**vbracket R package:**
- Location: `/Users/yoshiakisato/Desktop/github/vbracket_github`
- Repository: https://github.com/yoshiakisato/vbracket
- Purpose: Custom legends with vertical statistical comparison brackets
- Used by: `sato_line_grouped_error_raw()` for 3+ groups

## Version History & Major Changes

### Recent Major Changes (This Session)

**1. R Code Reproduction Fix**
- Modified `extractRelevantRCode()` to use `window.lastPlotSettings`
- Added settings storage in `renderPlotPngFromSelectionFixed()` (lines 9892-9909)
- Ensures R code output reproduces exact preview figure

**2. VBracket Integration**
- Implemented vbracket legend for 3+ groups line plots
- Added UI controls in Statistics tab (lines 747-794)
- Modified `sato_line_grouped_error_raw()` to parse statistics and create legend
- Supports both preset positions and custom X/Y coordinates

**3. Custom X/Y Positioning**
- Added "Custom (X/Y)" option to vbracket position selector
- Added X and Y position number inputs (0-1 scale)
- Modified R code to use `legend_x`/`legend_y` when position="custom"

**4. UI Organization**
- Moved "Export statistics to selected cells" button from Statistics tab to Data & Chart tab
- Adjusted button width and font size for consistency
- Verified no duplicate controls across tabs

### Known Issues

**1. VBracket Package Installation**
- vbracket is not automatically installed in webR
- Users must install manually or code gracefully handles missing package
- Future: Consider bundling or auto-installation

**2. Large File Size**
- taskpane.js is 473KB (10,000+ lines)
- All R functions embedded as strings
- Future: Consider splitting into multiple files or loading R code separately

**3. Font Availability**
- Some fonts may not render in svglite/webR
- Cross-platform consistency not guaranteed
- Future: Font subsetting or web font loading

**4. Browser Compatibility**
- Tested primarily in Excel for Windows/Mac desktop
- Web version of Excel may have limitations
- WebAssembly (webR) requires modern browser

## Contact & Contributions

**Author:** Yoshiaki Sato
**ORCID:** 0000-0003-3375-5189

**Related Repositories:**
- vbracket package: (Repository URL if public)

**Feedback:** Report issues or feature requests to project maintainer.

---

*This document is a living memory file for Claude Code sessions. Update as major features are added or architectural decisions are made.*
