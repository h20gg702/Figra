# Figra

**Publication-Quality Scientific Figures from Excel Data**

Figra is an Excel Office Add-in that creates publication-ready scientific visualizations using R and ggplot2 — directly in your browser, with no R installation required.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Excel%20(Windows%20%7C%20Mac%20%7C%20Web)-green.svg)

## Features

### Chart Types (21+)
- **Distribution**: Histogram, Box plot, Violin plot, Dot plot
- **Bar charts**: Simple, with error bars, with data points
- **Grouped charts**: Grouped bar, box, violin (with/without dots)
- **Line plots**: Simple, grouped, with error ribbons
- **Dose-response**: 4-Parameter Logistic (4PL) IC50 curves

## Chart Examples

### Distribution Charts

| Histogram | Box Plot | Box Plot + Dots |
|:---------:|:--------:|:---------------:|
| ![Histogram](docs/images/histogram.png) | ![Box Plot](docs/images/box.png) | ![Box Plot + Dots](docs/images/box_dot.png) |

| Violin Plot | Violin Plot + Dots | Dot Plot |
|:-----------:|:------------------:|:--------:|
| ![Violin Plot](docs/images/violin.png) | ![Violin Plot + Dots](docs/images/violin_dot.png) | ![Dot Plot](docs/images/dot.png) |

### Bar Charts

| Bar Plot | Bar + Error | Bar + Error + Dots |
|:--------:|:-----------:|:------------------:|
| ![Bar Plot](docs/images/bar.png) | ![Bar + Error](docs/images/bar_error.png) | ![Bar + Error + Dots](docs/images/bar_error_dot.png) |

### Grouped Bar Charts

| Grouped Bar | Grouped Bar + Error | Grouped Bar + Error + Dots |
|:-----------:|:-------------------:|:--------------------------:|
| ![Grouped Bar](docs/images/bar_grouped.png) | ![Grouped Bar + Error](docs/images/bar_grouped_error.png) | ![Grouped Bar + Error + Dots](docs/images/bar_grouped_error_dot.png) |

### Grouped Box & Violin Charts

| Grouped Box | Grouped Box + Dots | Grouped Violin + Dots |
|:-----------:|:------------------:|:---------------------:|
| ![Grouped Box](docs/images/box_grouped.png) | ![Grouped Box + Dots](docs/images/box_grouped_dot.png) | ![Grouped Violin + Dots](docs/images/violin_grouped_dot.png) |

### Line Charts

| Line Plot | Grouped Line | Grouped Line + Error |
|:---------:|:------------:|:--------------------:|
| ![Line Plot](docs/images/line.png) | ![Grouped Line](docs/images/line_grouped.png) | ![Grouped Line + Error](docs/images/line_grouped_error.png) |

| Grouped Line + Error (Raw Data) |
|:-------------------------------:|
| ![Grouped Line + Error Raw](docs/images/line_grouped_error_raw.png) |

### Dose-Response Curve

| IC50 Dose-Response (4PL) |
|:------------------------:|
| ![IC50 Dose-Response](docs/images/ic50_dose_response.png) |

### Statistical Analysis Examples

| Pairwise Comparisons | VBracket Legend (3+ groups) |
|:--------------------:|:---------------------------:|
| ![Stats Brackets](docs/images/stats_brackets.png) | ![VBracket Legend](docs/images/vbracket_legend.png) |

### Statistical Analysis
- **Automatic test selection** based on data normality and group count
- **Parametric tests**: t-test, Welch's t-test, ANOVA
- **Non-parametric tests**: Wilcoxon, Kruskal-Wallis
- **Post-hoc tests**: Tukey HSD, Bonferroni, Holm, Dunnett, Dunn
- **Variance tests**: Levene, F-test
- **Display options**: Stars (*, **, ***), letters (a, b, c), or p-values

### Customization
- **Themes**: Minimal, Classic, BW, Light, Dark, and more
- **Fonts**: Arial, Times, Helvetica, and system fonts
- **Colors**: Full color customization with opacity control
- **Text formatting**: Italic, superscript, subscript support
- **Axis scales**: Linear, log10, log2, ln, -log10, -log2

### Export & Reproducibility
- **Insert figures** directly into Excel sheets
- **Download R code** that reproduces the exact figure
- **Export statistics** to Excel cells
- **Save/load presets** for consistent styling

## How It Works

Figra uses [webR](https://webr.r-wasm.org/) to run R code directly in your browser via WebAssembly. This means:

- **No R installation required** on your computer
- **No server needed** — all processing happens locally
- **Cross-platform** — works on Windows, Mac, and Excel Online
- **Reproducible** — export R code to recreate figures anywhere

## Installation

### Option 1: Sideload for Development

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/Figra.git
   cd Figra/addin
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm run dev-server
   ```

4. Sideload the add-in in Excel:
   - Open Excel
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - Click **Upload My Add-in**
   - Select `manifest/manifest.xml`

### Option 2: Use Hosted Version

*(Coming soon)*

## Quick Start

1. **Select your data** in Excel (with headers)
2. **Click "Load Data"** in the Figra panel
3. **Choose a chart type** from the visual dropdown
4. **Map your columns** (Group, X, Y, Error as needed)
5. **Customize** colors, fonts, and themes
6. **Click "Preview"** to see your figure
7. **Click "Insert Figure"** to add it to your sheet

## Supported Data Formats

| Chart Type | Required Columns |
|------------|------------------|
| Histogram | Values |
| Box/Violin/Dot | Category, Values |
| Bar + Error | Category, Mean, Error |
| Grouped charts | Group, Category, Values |
| Line + Error (raw) | Group, X, Values (calculates stats) |
| IC50 Dose-Response | Concentration, Response |

## Statistical Analysis

Figra automatically performs appropriate statistical tests:

| Groups | Normal Data | Non-normal Data |
|--------|-------------|-----------------|
| 2 | t-test / Welch's | Wilcoxon |
| 3+ | ANOVA + post-hoc | Kruskal-Wallis + Dunn |

Results are displayed as:
- **Brackets with symbols** on the figure
- **VBracket legend** for grouped line plots (3+ groups)
- **Exportable statistics** to Excel cells

## Project Structure

```
Figra/
├── addin/                 # Excel Add-in source code
│   ├── src/
│   │   ├── taskpane/     # Main UI and logic
│   │   └── commands/     # Office commands
│   ├── assets/           # Icons and images
│   ├── package.json
│   └── webpack.config.js
├── manifest/
│   └── manifest.xml      # Office Add-in manifest
├── docs/                  # Documentation
└── scripts/              # Build scripts
```

## Technologies

- **Office.js** — Excel JavaScript API
- **webR** — R compiled to WebAssembly
- **ggplot2** — Grammar of graphics for R
- **svglite** — SVG output from R
- **canvg** — SVG to Canvas conversion

## Development

### Prerequisites
- Node.js 16+
- npm or yarn

### Commands
```bash
npm run build        # Production build
npm run build:dev    # Development build
npm run dev-server   # Start dev server with hot reload
npm run lint         # Run linter
npm run validate     # Validate manifest
```

### Adding New Chart Types

1. Add the chart type to the dropdown in `taskpane.html`
2. Create the `sato_charttype()` R function in `taskpane.js`
3. Add educational R code generation in `generateEducationalRCode()`
4. Update column mapping logic if needed

## Author

**Yoshiaki Sato**
ORCID: [0000-0003-3375-5189](https://orcid.org/0000-0003-3375-5189)

## License

MIT License — see [LICENSE](LICENSE) for details.

## Acknowledgments

- [webR](https://webr.r-wasm.org/) by George Stagg and the R Consortium
- [ggplot2](https://ggplot2.tidyverse.org/) by Hadley Wickham
- [Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/) documentation

---

**Figra** — From Data to Publication-Ready Figures
