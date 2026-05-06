# Figra

**Publication-Quality Scientific Figures from Excel Data**

Figra is an Excel Office Add-in that creates publication-ready scientific visualizations using R and ggplot2 — directly in your browser, with no R installation required.

![License](https://img.shields.io/badge/license-Custom%20Terms-blue.svg)
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

| Histogram | Box Plot | Violin Plot |
|:---------:|:--------:|:-----------:|
| ![Histogram](https://h20gg702.github.io/figra-pages/figures/histogram.png) | ![Box Plot](https://h20gg702.github.io/figra-pages/figures/box.png) | ![Violin Plot](https://h20gg702.github.io/figra-pages/figures/violin.png) |

### Bar & Grouped Charts

| Bar + Error + Dots | Grouped Bar + Error + Dots | Grouped Box + Dots |
|:------------------:|:--------------------------:|:------------------:|
| ![Bar + Error + Dots](https://h20gg702.github.io/figra-pages/figures/bar.png) | ![Grouped Bar + Error + Dots](https://h20gg702.github.io/figra-pages/figures/grouped_bar_error_dots.png) | ![Grouped Box + Dots](https://h20gg702.github.io/figra-pages/figures/grouped_box_dots.png) |

### Line & Dose-Response Charts

| Grouped Line | Grouped Dose-Response | IC50 Dose-Response (4PL) |
|:------------:|:---------------------:|:------------------------:|
| ![Grouped Line](https://h20gg702.github.io/figra-pages/figures/group_line.png) | ![Grouped Dose-Response](https://h20gg702.github.io/figra-pages/figures/group_dose.png) | ![IC50 Dose-Response](https://h20gg702.github.io/figra-pages/figures/dose_response_curve.png) |

### Specialized Charts

| Scatter Plot | Survival Curve (LQ) | Gene Ontology |
|:------------:|:-------------------:|:-------------:|
| ![Scatter Plot](https://h20gg702.github.io/figra-pages/figures/scatter_plot.png) | ![Survival Curve](https://h20gg702.github.io/figra-pages/figures/lq.png) | ![Gene Ontology](https://h20gg702.github.io/figra-pages/figures/gene_ontology.png) |

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
   git clone https://github.com/h20gg702/Figra.git
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

### Option 2: Install from Microsoft AppSource

Search for **Figra** in Microsoft AppSource, or install directly from Excel via **Insert → Add-ins → Get Add-ins**.

The hosted add-in is served from [https://h20gg702.github.io/figra-pages/](https://h20gg702.github.io/figra-pages/).

## Get Started in 3 Steps

**Step 1 — Download the manifest file**
Download `manifest.xml` from [https://h20gg702.github.io/figra-pages/](https://h20gg702.github.io/figra-pages/)

**Step 2 — Add to Excel**

🌐 **Excel Online (easiest — any OS):**
1. Open a workbook at office.com
2. **Home** → **Add-ins** → **More Settings** → **Upload My Add-in**
3. Select `manifest.xml` → **Upload**

🪟 **Windows Desktop Excel:**
1. Create a folder e.g. `C:\FigraAddin` and copy `manifest.xml` into it
2. Right-click the folder → **Properties** → **Sharing** tab → **Advanced Sharing** → check **Share this folder** → note the network path (e.g. `\\YourPC\FigraAddin`)
3. In Excel: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**
4. Paste the network path → **Add Catalog** → check **Show in Menu** → **OK**
5. Close and reopen Excel → **Insert** → **Get Add-ins** → **SHARED FOLDER** tab → select **Figra** → **Add**

> ⚠️ Microsoft disabled Shared Folder sideloading by default from Office build 16.0.16227 (March 2023). If the Shared Folder tab is missing, use Excel Online or ask your IT admin to deploy via Microsoft 365 Admin Center.

🍎 **Mac Desktop Excel:**
1. In Finder press **Cmd+Shift+G** → navigate to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create the `wef` folder if it does not exist)
2. Copy `manifest.xml` into this folder
3. Open Excel → **Home** → **Add-ins** → select **Figra**

**Step 3 — Start creating figures**
Click **Figra** in the Home tab, register (one-time), and create your first figure!

---

## Frequently Asked Questions

**Is Figra free?**
Yes, completely free with a one-time registration.

**Do I need to install R?**
No. Figra uses [webR](https://webr.r-wasm.org/) — R runs directly in your browser via WebAssembly with no installation required.

**Does it work on Mac?**
Yes. Excel Online provides full functionality on Mac. The Mac Excel desktop app has some limitations due to OS restrictions: saving PNG with embedded metadata and downloading `.R` files directly don't work, but you can use Insert Figure to Sheet and Copy R Code instead.

**What file formats can I export?**
PNG with embedded metadata (data and settings), allowing you to regenerate and re-edit figures later.

**What statistical tests are supported?**
t-test, Wilcoxon, ANOVA, and Kruskal-Wallis. Post-hoc tests: Tukey HSD, Dunnett, Steel, and Dunn, with Bonferroni and Holm correction options.

**What kind of data works best with Figra?**
Figra runs R inside your browser via WebAssembly, which can use up to 4GB of your PC's memory ([details](https://v8.dev/blog/4gb-wasm-memory)). Performance depends on your available RAM and CPU. Typical experimental datasets work well on most modern computers, but very large datasets may be slow. Figra is not designed for large-scale genomics data visualization such as RNA-seq or other NGS results.

---

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

Free for academic, research, and personal use. Redistribution, modification, and commercial resale are not permitted without written permission from the author. See [LICENSE](LICENSE) and the full [Terms of Use](https://h20gg702.github.io/figra-pages/terms.html) for details.

## Acknowledgments

- [webR](https://webr.r-wasm.org/) by George Stagg and the R Consortium
- [ggplot2](https://ggplot2.tidyverse.org/) by Hadley Wickham
- [Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/) documentation

---

**Figra** — From Data to Publication-Ready Figures
