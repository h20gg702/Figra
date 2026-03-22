# Statistical Analysis Refactoring Plan

## Goal
Extract duplicated per-chart-type stat logic into shared R helper functions, so any new chart type gets correct statistics automatically and fixes apply everywhere.

---

## Current Architecture (Two Separate Systems)

### System 1: Visual Bracket Annotations on Plot
**Function:** `sato_add_statistics_to_plot()`
**Location:** taskpane.js ~line 13168
**Purpose:** Adds visual bracket lines + significance symbols ON the ggplot (using ggpubr)
**Used by:** box, box_dot, violin_dot, bar_error_dot, box_grouped, box_grouped_dot, violin_grouped_dot, bar_grouped_error_dot
**NOT used by:** line_grouped_error_raw, lq_survival_grouped (these have inline blocks)

### System 2: Stat Text Export (for "Export Stats" / display panel)
**Purpose:** Generates formatted text with normality/variance/test/post-hoc results for export to Excel cells
**Problem:** Implemented separately (inline) in each R function — NOT shared

---

## Where Stat Text Logic Lives (Per Chart Type)

**ALL chart types have their own copy-pasted inline stat block. NONE share a function.**
They all write to the same global variable name but the stat logic is duplicated each time.

| Chart Type | Inline Stat Block Lines | Global R Variable | Paired Support | Style | Notes |
|---|---|---|---|---|---|
| `box_grouped_dot` | ~17206–18277 | `grouped_bar_stat_results` | ✅ Yes | Verbose | Reference implementation |
| `bar_grouped_error_dot` (v1) | ~18405–18613 | `grouped_bar_stat_results` | ❌ **Missing** | Condensed | variance text: "equal"/"unequal" |
| `violin_grouped_dot` | ~18782–19855 | `grouped_bar_stat_results` | ✅ Yes | Verbose | Nearly identical to box_grouped_dot |
| `bar_grouped_error_dot` (v2) | ~20060–20162 | `grouped_bar_stat_results` | ❌ **Missing** | Minified | same logic as v1, just minified |
| `bar_grouped_error_dot` (v3) | ~20329–21402 | `grouped_bar_stat_results` | ✅ Yes | Verbose | Nearly identical to box_grouped_dot |
| `line_grouped_error_raw` | ~15804–16553 | `line_plot_stat_results` | ✅ Yes | Verbose | per-x-value loop; vbracket 3+ groups |
| `lq_survival_grouped` | ~23130–23558 | `lq_stat_results` | ✅ Yes | Verbose | per-dose loop; log10 symbols; vbracket 3+ groups |

**Known differences between versions (bugs to fix during refactor):**
1. **Paired test missing in bar_grouped_error_dot v1 and v2** — `bar_grouped_error_dot` does not handle paired data at all
2. **Variance text wording**: v1/v2 say "equal"/"unequal" — should be "equal variances"/"unequal variances"
3. **3+-group auto string**: v1/v2 use `"anova"`/`"kruskal"` — should be `"t-test"`/`"wilcoxon"` style (cosmetic, no functional impact)

> Note: `sato_add_statistics_to_plot()` (line ~13168) is a SEPARATE function that handles
> visual ggpubr bracket drawing on the plot for grouped chart types. It is NOT the stat text
> export block. Both exist side by side in the grouped chart functions.

---

## Confirmed Statistical Workflow Spec

### 2-Group
- **Auto**: Shapiro-Wilk per group → both normal: variance test → Student's or Welch's t-test; any non-normal: Wilcoxon
- **Manual parametric**: skip normality → variance test → Student's or Welch's t-test
- **Manual non-parametric**: skip normality → Wilcoxon
- **Paired** (any mode): same but paired t-test / paired Wilcoxon; NO variance test
- **Variance test** (unpaired + parametric only): Levene default / F-test if user selects; p > 0.05 → equal → Student's; p ≤ 0.05 → unequal → Welch's

### 3+-Group Unpaired
- **Auto**: Shapiro-Wilk per group → all normal: ANOVA; any non-normal: Kruskal-Wallis
- **Manual parametric**: ANOVA; **Manual non-parametric**: Kruskal-Wallis
- **Post-hoc** (only when omnibus p < 0.05):
  - ANOVA → **Tukey HSD** (default) / Bonferroni / Holm / **Dunnett** (vs control group)
  - Kruskal-Wallis → **Dunn + Bonferroni** (default) / **Steel** (vs control group, non-parametric Dunnett equivalent)

### 3+-Group Paired
- **Auto**: Shapiro-Wilk on pairwise differences → all normal: RM-ANOVA; any non-normal: Friedman
- **Manual parametric**: RM-ANOVA; **Manual non-parametric**: Friedman
- **Post-hoc** (p < 0.05 only): paired pairwise t-test (RM-ANOVA) / paired pairwise Wilcoxon (Friedman)

---

## Architecture: Separation of Computation and Display

The shared helper does **computation only** — returns structured results.
Each chart type decides how to **display** those results.

### Helper returns:
```r
list(
  result_text,    # formatted string for text export to Excel
  sig_label,      # "***", "**", "*", "ns" — for visual annotation
  p_val,          # raw p-value
  test_name,      # "Student's t-test", "Welch's t-test", "Wilcoxon", etc.
  group_pairs,    # list of (group1, group2) pairs with p-values — for bracket drawing
  omnibus_p       # 3+-group: overall test p-value
)
```

### Display methods (caller's responsibility):
1. **Default ggpubr brackets** — box_grouped_dot, violin_grouped_dot, bar_grouped_error_dot: pass `group_pairs` → `bracket_data` → `stat_pvalue_manual()`
2. **VBracket legend** — line_grouped_error_raw, lq_survival_grouped (3+ groups): parse `result_text` or `group_pairs` → `legend_bracket()`
3. **Log10 symbols below bars** — lq_survival_grouped (2-group): place `sig_label` at computed y position below error bars
4. **Text export** — all chart types: store `result_text` → global R variable → JS retrieves for Excel export

---

## Standard Stat Flow (the pattern we want to share)

All Group A chart types implement this same flow:

### 2-Group Flow
```
1. If paired: extract subject_id, merge, align pairs
2. Normality:
   - Paired: shapiro.test(differences), n >= 3 && n <= 5000
   - Unpaired: shapiro.test(group1), shapiro.test(group2), each n >= 3 && n <= 5000
3. Auto test selection:
   - "auto" + both_normal → "t-test"
   - "auto" + !both_normal → "wilcoxon"
   - "parametric" → "t-test"
   - "nonparametric" → "wilcoxon"
4. Variance test (unpaired + parametric only):
   - variance_test == "levene" → anova(lm(abs_deviations ~ group)), threshold p > 0.05
   - variance_test == "f-test" → var.test(), threshold p > 0.05
5. Run test:
   - Paired Wilcoxon: wilcox.test(paired=TRUE)
   - Paired t-test: t.test(paired=TRUE)
   - Wilcoxon: wilcox.test()
   - t-test: t.test(var.equal=equal_variances)
     → equal: "Student's t-test", unequal: "Welch's t-test"
6. Format result_text:
   ">> X-axis value: N\nGroups: A vs B\nSummary...\nNormality...\nVariance...\nTest: X, p=Y (sig)"
```

### 3+-Group Flow (Unpaired)
```
1. Build anova_data: data.frame(group=factor(...), value=numeric(...))
2. Normality per group: shapiro.test(grp_data), n >= 3 && n <= 5000
3. Decide: use_nonparametric = (statistical_test == "nonparametric") || (statistical_test == "auto" && !all_normal)
4. Non-parametric path:
   - kruskal.test(value ~ group)
   - If p < 0.05: post-hoc based on post_hoc_test:
     - "steel" → Steel.test() via kSamples
     - "dunn_holm" → dunn.test(method="holm")
     - default → dunn.test(method="bonferroni")
5. Parametric path:
   - aov(value ~ group)
   - If p < 0.05: post-hoc based on post_hoc_test:
     - "bonferroni" → pairwise.t.test(p.adjust.method="bonferroni")
     - "holm" → pairwise.t.test(p.adjust.method="holm")
     - default (tukey) → TukeyHSD()
6. Format result_text with Overall test + Post-hoc lines
```

### 3+-Group Flow (Paired)
```
1. Build anova_data with subject_id column
2. Normality of pairwise differences: for each pair, merge by subject_id, shapiro.test(diffs)
3. Auto: all_normal → "rm_anova", else → "friedman"
4. RM-ANOVA: aov(value ~ group + Error(subject_id/group))
5. Friedman: friedman.test(value ~ group | subject_id)
6. Post-hoc (if omnibus p < 0.05):
   - RM-ANOVA: pairwise.t.test(paired=TRUE, p.adjust.method=ph_method)
   - Friedman: pairwise.wilcox.test(paired=TRUE, p.adjust.method=ph_method)
```

---

## Proposed Refactoring

### New Shared R Helper: `sato_run_stats_2group()`
```r
sato_run_stats_2group <- function(
  group1_data, group2_data, group_names,
  x_label,           # label for ">> X-axis value: ..." line
  statistical_test, variance_test,
  paired = FALSE,
  error_label = "SD"  # for summary line
) -> list(result_text, sig_label, p_val)
```

### New Shared R Helper: `sato_run_stats_ngroup()`
```r
sato_run_stats_ngroup <- function(
  anova_data,         # data.frame(group, value, [subject_id])
  x_label,
  statistical_test, post_hoc_test, dunnett_control,
  paired = FALSE,
  error_label = "SD"
) -> list(result_text, omnibus_p)
```

Both helpers:
- Take data + parameters
- Run the standard flow
- Return `result_text` (formatted string) + p-values
- Called from ALL chart types including line_grouped_error_raw and lq_survival_grouped

### Where to define them
Define as a JS constant string at the top of the R code section, prepended to every R execution.
Example:
```javascript
const SHARED_STAT_HELPERS_R = `
sato_run_stats_2group <- function(...) { ... }
sato_run_stats_ngroup <- function(...) { ... }
`;
```
Then every R code block starts with `${SHARED_STAT_HELPERS_R}`.

---

## Code-Level Verification Findings

### 1. `get_sig_symbol()` — locally defined 8 times
Defined inside every chart type's R block separately. It uses `stat_symbol_type` from the enclosing function scope.
The shared helper must **include `get_sig_symbol()` as an internal function** (taking `stat_symbol_type` as a parameter).

### 2. Parameter name inconsistency — post-hoc test
| Chart type | Parameter name |
|---|---|
| `line_grouped_error_raw` | `post_hoc_test` |
| `lq_survival_grouped` | `post_hoc_test` |
| `box_grouped_dot`, `violin_grouped_dot`, `bar_grouped_error_dot` | `selected_posthoc_test` |

The shared helper will use **`post_hoc_test`** as the canonical name. Callers using `selected_posthoc_test` pass it as `post_hoc_test=selected_posthoc_test`.

### 3. `error_type` — available in all grouped chart scopes
All grouped chart functions receive `error_type` ("sd", "se", "ci95") as a function parameter. Used in summary stats output. Consistent across all.

### 4. `stat_symbol_type` — available in all scopes
Passed as a function parameter everywhere. Consistent.

### 5. There is already a `sato_perform_statistical_test()` (~line 13024)
A different existing shared function — handles single-value stat test execution. NOT the same as what we're building. Our new helper handles the full flow (normality → variance → test → post-hoc → formatted text).

### 6. How R code is assembled
Each chart type builds a complete R string as a JS template literal and sends it to webR independently. The shared helper R string must be **prepended inside each chart type's R execution block** — there is no single global R execution entry point.

### 7. `dunnett_control` — control group for Dunnett/Steel tests
Available in all chart type scopes. Must be passed to the helper.

### 8. `paired` and `subject_id` column
- Grouped charts: `paired` boolean + data has `subject_id` column when paired
- LQ: uses `subject_col` index to locate subject column in raw data (different approach)
- Helper must handle both approaches — accept pre-merged paired data

---

## Complete Helper R Code (ready to paste into taskpane.js)

Source of truth: extracted from `line_grouped_error_raw` stat block (the reference implementation).
In `taskpane.js`, newlines inside R strings must be `\\n` (escaped). Below uses real newlines for readability —
when adding to the JS constant, replace all `\n` inside the R string with `\\n`.

### JS constant to add near top of R execution section

```javascript
const SHARED_STAT_HELPERS_R = `
sato_run_stats_2group <- function(
  group1_data, group2_data, group_names,
  x_label,
  statistical_test, variance_test,
  stat_symbol_type = "stars",
  paired = FALSE,
  error_label = "SD",
  desc_lines = c()
) {
  get_sig_sym <- function(p) {
    if (stat_symbol_type == "pvalue") return(sprintf("p=%.3f", p))
    if (p < 0.001) "***" else if (p < 0.01) "**" else if (p < 0.05) "*" else "ns"
  }

  normality_text <- c()
  both_normal <- TRUE

  if (paired) {
    diffs <- group1_data - group2_data
    if (length(diffs) >= 3 && length(diffs) <= 5000) {
      sw <- tryCatch(shapiro.test(diffs), error = function(e) NULL)
      if (!is.null(sw)) {
        both_normal <- sw$p.value >= 0.05
        normality_text <- c(normality_text,
          sprintf("  Differences: p=%.4f (%s)", sw$p.value, if (both_normal) "normal" else "non-normal"))
      }
    }
  } else {
    is_n1 <- TRUE; is_n2 <- TRUE
    if (length(group1_data) >= 3 && length(group1_data) <= 5000) {
      sw <- tryCatch(shapiro.test(group1_data), error = function(e) NULL)
      if (!is.null(sw)) {
        is_n1 <- sw$p.value >= 0.05
        normality_text <- c(normality_text,
          sprintf("  %s: p=%.4f (%s)", group_names[1], sw$p.value, if (is_n1) "normal" else "non-normal"))
      }
    }
    if (length(group2_data) >= 3 && length(group2_data) <= 5000) {
      sw <- tryCatch(shapiro.test(group2_data), error = function(e) NULL)
      if (!is.null(sw)) {
        is_n2 <- sw$p.value >= 0.05
        normality_text <- c(normality_text,
          sprintf("  %s: p=%.4f (%s)", group_names[2], sw$p.value, if (is_n2) "normal" else "non-normal"))
      }
    }
    both_normal <- is_n1 && is_n2
  }

  test_to_use <- statistical_test
  if (statistical_test == "auto") {
    test_to_use <- if (both_normal) "t-test" else "wilcoxon"
  } else if (statistical_test == "parametric") {
    test_to_use <- "t-test"
  } else if (statistical_test == "nonparametric") {
    test_to_use <- "wilcoxon"
  }

  variance_text <- ""
  equal_variances <- TRUE
  if (test_to_use != "wilcoxon" && !paired) {
    if (variance_test == "levene") {
      cd <- data.frame(
        values = c(group1_data, group2_data),
        group  = factor(c(rep(group_names[1], length(group1_data)),
                          rep(group_names[2], length(group2_data))))
      )
      gm <- tapply(cd$values, cd$group, mean)
      ad <- abs(cd$values - gm[cd$group])
      lv <- tryCatch(anova(lm(ad ~ cd$group)), error = function(e) NULL)
      if (!is.null(lv)) {
        lv_p <- lv$\`Pr(>F)\`[1]
        equal_variances <- lv_p > 0.05
        variance_text <- sprintf("Variance test: p=%.4f (%s, Levene)", lv_p,
                                 if (equal_variances) "equal variances" else "unequal variances")
      }
    } else {
      vt <- tryCatch(var.test(group1_data, group2_data), error = function(e) NULL)
      if (!is.null(vt)) {
        equal_variances <- vt$p.value > 0.05
        variance_text <- sprintf("Variance test: p=%.4f (%s, F-test)", vt$p.value,
                                 if (equal_variances) "equal variances" else "unequal variances")
      }
    }
  }

  test_result <- NULL; test_name <- ""
  if (paired) {
    if (test_to_use == "wilcoxon") {
      test_result <- tryCatch(wilcox.test(group1_data, group2_data, paired = TRUE), error = function(e) NULL)
      test_name <- "Paired Wilcoxon"
    } else {
      test_result <- tryCatch(t.test(group1_data, group2_data, paired = TRUE), error = function(e) NULL)
      test_name <- "Paired t-test"
    }
  } else if (test_to_use == "wilcoxon") {
    test_result <- tryCatch(wilcox.test(group1_data, group2_data), error = function(e) NULL)
    test_name <- "Wilcoxon"
  } else {
    test_result <- tryCatch(t.test(group1_data, group2_data, var.equal = equal_variances), error = function(e) NULL)
    test_name <- if (equal_variances) "Student's t-test" else "Welch's t-test"
  }

  if (is.null(test_result)) return(list(result_text = "", sig_label = "ns", p_val = NA))

  p_val <- test_result$p.value
  sig_label <- get_sig_sym(p_val)

  result_text <- sprintf(">> X-axis value: %s\nGroups: %s vs %s", x_label, group_names[1], group_names[2])
  if (length(desc_lines) > 0)
    result_text <- paste0(result_text, sprintf("\nSummary (mean +/- %s):\n", error_label), paste(desc_lines, collapse="\n"))
  if (length(normality_text) > 0)
    result_text <- paste0(result_text, "\nNormality (Shapiro-Wilk):\n", paste(normality_text, collapse="\n"))
  else
    result_text <- paste0(result_text, "\nNormality (Shapiro-Wilk): skipped (n < 3 per group)")
  if (nchar(variance_text) > 0)
    result_text <- paste0(result_text, "\n", variance_text)
  result_text <- paste0(result_text, sprintf("\nTest: %s, p=%.4f (%s)", test_name, p_val, sig_label))

  list(result_text = result_text, sig_label = sig_label, p_val = p_val)
}

sato_run_stats_ngroup <- function(
  anova_data,        # data.frame(group=factor, value=numeric, [subject_id=factor for paired])
  x_label,
  statistical_test, post_hoc_test, dunnett_control = "",
  stat_symbol_type = "stars",
  paired = FALSE,
  error_label = "SD",
  desc_lines = c()
) {
  get_sig_sym <- function(p) {
    if (stat_symbol_type == "pvalue") return(sprintf("p=%.3f", p))
    if (p < 0.001) "***" else if (p < 0.01) "**" else if (p < 0.05) "*" else "ns"
  }

  result_text <- sprintf(">> X-axis value: %s\nTime point %s:", x_label, x_label)
  if (length(desc_lines) > 0)
    result_text <- paste0(result_text, sprintf("\nSummary (mean +/- %s):\n", error_label), paste(desc_lines, collapse="\n"))

  if (paired) {
    # --- PAIRED: RM-ANOVA or Friedman ---
    normality_text <- c(); all_normal <- TRUE
    groups_list <- levels(anova_data$group)
    for (pair_n in combn(groups_list, 2, simplify = FALSE)) {
      g1 <- anova_data[anova_data$group == pair_n[1], c("subject_id","value")]
      g2 <- anova_data[anova_data$group == pair_n[2], c("subject_id","value")]
      mg <- merge(g1, g2, by = "subject_id")
      diffs_n <- mg$value.x - mg$value.y
      if (length(diffs_n) >= 3 && length(diffs_n) <= 5000) {
        sw <- tryCatch(shapiro.test(diffs_n), error = function(e) NULL)
        if (!is.null(sw)) {
          is_norm <- sw$p.value >= 0.05
          normality_text <- c(normality_text,
            sprintf("  %s vs %s: p=%.4f (%s)", pair_n[1], pair_n[2], sw$p.value,
                    if (is_norm) "normal" else "non-normal"))
          if (!is_norm) all_normal <- FALSE
        }
      }
    }
    if (length(normality_text) > 0)
      result_text <- paste0(result_text, "\nNormality of differences (Shapiro-Wilk):\n", paste(normality_text, collapse="\n"))

    test_to_use_p <- statistical_test
    if (statistical_test == "auto")       test_to_use_p <- if (all_normal) "rm_anova" else "friedman"
    else if (statistical_test == "parametric")  test_to_use_p <- "rm_anova"
    else                                         test_to_use_p <- "friedman"

    omnibus_p <- NA; test_name_p <- ""
    if (test_to_use_p == "rm_anova") {
      rm_res <- tryCatch(summary(aov(value ~ group + Error(subject_id/group), data = anova_data)), error = function(e) NULL)
      if (!is.null(rm_res)) {
        omnibus_p <- tryCatch(rm_res[["Error: subject_id:group"]][[1]][["Pr(>F)"]][1], error = function(e) NA)
        if (is.null(omnibus_p) || length(omnibus_p) == 0) omnibus_p <- NA_real_
        test_name_p <- "RM-ANOVA"
      }
    } else {
      fr <- tryCatch(friedman.test(value ~ group | subject_id, data = anova_data), error = function(e) NULL)
      if (!is.null(fr)) { omnibus_p <- fr$p.value; test_name_p <- "Friedman" }
    }

    if (!is.na(omnibus_p)) {
      result_text <- paste0(result_text,
        sprintf("\nOverall test: %s, p=%.4f (%s)", test_name_p, omnibus_p, get_sig_sym(omnibus_p)))
      if (omnibus_p < 0.05) {
        ph_method <- if (post_hoc_test %in% c("bonferroni","holm")) post_hoc_test else "holm"
        ph_label  <- if (ph_method == "bonferroni") "Bonferroni" else "Holm"
        ph_func   <- if (test_to_use_p == "rm_anova") pairwise.t.test else pairwise.wilcox.test
        ph_name   <- if (test_to_use_p == "rm_anova") "Pairwise paired t-test" else "Pairwise paired Wilcoxon"
        ph_res <- tryCatch(ph_func(anova_data$value, anova_data$group, p.adjust.method=ph_method, paired=TRUE), error=function(e) NULL)
        if (!is.null(ph_res)) {
          result_text <- paste0(result_text, sprintf("\nPost-hoc (%s, %s):", ph_name, ph_label))
          p_mat <- ph_res$p.value
          for (r in rownames(p_mat)) for (c in colnames(p_mat)) {
            p_adj <- p_mat[r,c]
            if (!is.na(p_adj)) result_text <- paste0(result_text,
              sprintf("\n  %s-%s: p=%.4f (%s)", r, c, p_adj, get_sig_sym(p_adj)))
          }
        }
      }
    }
    return(list(result_text = result_text, omnibus_p = omnibus_p))
  }

  # --- UNPAIRED: ANOVA or Kruskal-Wallis ---
  normality_text <- c(); all_normal <- TRUE
  for (grp in levels(anova_data$group)) {
    gd <- anova_data$value[anova_data$group == grp]
    if (length(gd) >= 3 && length(gd) <= 5000) {
      sw <- tryCatch(shapiro.test(gd), error = function(e) NULL)
      if (!is.null(sw)) {
        norm_status <- if (sw$p.value >= 0.05) "normal" else "non-normal"
        if (sw$p.value < 0.05) all_normal <- FALSE
        normality_text <- c(normality_text, sprintf("  %s: p=%.4f (%s)", grp, sw$p.value, norm_status))
      }
    }
  }
  if (length(normality_text) > 0)
    result_text <- paste0(result_text, "\nNormality (Shapiro-Wilk):\n", paste(normality_text, collapse="\n"))
  else
    result_text <- paste0(result_text, "\nNormality (Shapiro-Wilk): skipped (n < 3 per group)")

  use_nonparametric <- (statistical_test == "nonparametric") || (statistical_test == "auto" && !all_normal)
  omnibus_p <- NA

  if (use_nonparametric) {
    kw <- tryCatch(kruskal.test(value ~ group, data = anova_data), error = function(e) NULL)
    if (!is.null(kw)) {
      omnibus_p <- kw$p.value
      result_text <- paste0(result_text,
        sprintf("\nOverall test: Kruskal-Wallis, p=%.4f (%s)", omnibus_p, get_sig_sym(omnibus_p)))
      if (!is.na(omnibus_p) && omnibus_p < 0.05) {
        if (post_hoc_test == "steel") {
          ctrl <- if (nchar(dunnett_control) > 0) dunnett_control else levels(anova_data$group)[1]
          ok <- tryCatch({ if (!requireNamespace("kSamples",quietly=TRUE)) webr::install("kSamples"); library(kSamples); TRUE }, error=function(e) FALSE)
          if (ok) {
            result_text <- paste0(result_text, sprintf("\nPost-hoc (Steel test, vs %s):", ctrl))
            ctrl_vals <- anova_data$value[anova_data$group == ctrl]
            for (trt in levels(anova_data$group)[levels(anova_data$group) != ctrl]) {
              sr <- tryCatch(Steel.test(list(ctrl_vals, anova_data$value[anova_data$group == trt])), error=function(e) NULL)
              if (!is.null(sr)) result_text <- paste0(result_text,
                sprintf("\n  %s-%s: p=%.4f (%s)", ctrl, trt, sr$st[2], get_sig_sym(sr$st[2])))
            }
          } else result_text <- paste0(result_text, "\n[ERROR] Steel test requires kSamples package.")
        } else {
          dunn_method <- if (post_hoc_test == "dunn_holm") "holm" else "bonferroni"
          dunn_label  <- if (dunn_method == "holm") "Holm" else "Bonferroni"
          ok <- tryCatch({ if (!requireNamespace("dunn.test",quietly=TRUE)) webr::install("dunn.test"); library(dunn.test); TRUE }, error=function(e) FALSE)
          if (ok) {
            dr <- tryCatch(dunn.test(anova_data$value, anova_data$group, method=dunn_method), error=function(e) NULL)
            if (!is.null(dr)) {
              result_text <- paste0(result_text, sprintf("\nPost-hoc (Dunn test with %s):", dunn_label))
              for (i in seq_along(dr$comparisons)) {
                comp_clean <- gsub(" - ", "-", dr$comparisons[i])
                result_text <- paste0(result_text,
                  sprintf("\n  %s: p=%.4f (%s)", comp_clean, dr$P.adjusted[i], get_sig_sym(dr$P.adjusted[i])))
              }
            }
          } else result_text <- paste0(result_text, "\n[ERROR] Dunn test requires dunn.test package.")
        }
      }
    }
  } else {
    ar <- tryCatch(aov(value ~ group, data = anova_data), error = function(e) NULL)
    if (!is.null(ar)) {
      as_ <- summary(ar)
      anova_p <- as_[[1]][["Pr(>F)"]][1]
      omnibus_p <- anova_p
      result_text <- paste0(result_text,
        sprintf("\nOverall test: ANOVA, p=%.4f (%s)", anova_p, if (!is.na(anova_p)) get_sig_sym(anova_p) else ""))
      if (!is.na(anova_p) && anova_p < 0.05) {
        if (post_hoc_test == "dunnett") {
          ctrl <- if (nchar(dunnett_control) > 0) dunnett_control else levels(anova_data$group)[1]
          ok <- tryCatch({ if (!requireNamespace("multcomp",quietly=TRUE)) webr::install("multcomp"); library(multcomp); TRUE }, error=function(e) FALSE)
          if (ok) {
            dr <- tryCatch({
              anova_data$group <- relevel(anova_data$group, ref = ctrl)
              summary(glht(aov(value ~ group, data=anova_data), linfct=mcp(group="Dunnett")))
            }, error=function(e) NULL)
            if (!is.null(dr)) {
              result_text <- paste0(result_text, sprintf("\nPost-hoc (Dunnett, vs %s):", ctrl))
              pvals <- dr$test$pvalues; cnames <- names(dr$test$coefficients)
              for (i in seq_along(pvals)) result_text <- paste0(result_text,
                sprintf("\n  %s: p=%.4f (%s)", cnames[i], pvals[i], get_sig_sym(pvals[i])))
            }
          } else result_text <- paste0(result_text, "\n[ERROR] Dunnett requires multcomp package.")
        } else if (post_hoc_test %in% c("bonferroni","holm")) {
          ph_label <- if (post_hoc_test == "bonferroni") "Bonferroni" else "Holm"
          ph_res <- tryCatch(pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method=post_hoc_test), error=function(e) NULL)
          if (!is.null(ph_res)) {
            result_text <- paste0(result_text, sprintf("\nPost-hoc (Pairwise t-test with %s):", ph_label))
            p_mat <- ph_res$p.value
            for (r in rownames(p_mat)) for (c in colnames(p_mat)) {
              p_adj <- p_mat[r,c]
              if (!is.na(p_adj)) result_text <- paste0(result_text,
                sprintf("\n  %s-%s: p=%.4f (%s)", r, c, p_adj, get_sig_sym(p_adj)))
            }
          }
        } else {
          # Default: Tukey HSD
          tr <- tryCatch(TukeyHSD(ar), error=function(e) NULL)
          if (!is.null(tr)) {
            ts_ <- tr$group
            result_text <- paste0(result_text, "\nPost-hoc (Tukey HSD):")
            for (i in 1:nrow(ts_)) {
              comparison <- rownames(ts_)[i]
              p_adj <- ts_[i,"p adj"]; diff <- ts_[i,"diff"]
              result_text <- paste0(result_text,
                sprintf("\n  %s: diff=%.2f, p=%.4f (%s)", comparison, diff,
                        ifelse(is.na(p_adj),1.0,p_adj), if (!is.na(p_adj)) get_sig_sym(p_adj) else "ns"))
            }
          }
        }
      }
    }
  }
  list(result_text = result_text, omnibus_p = omnibus_p)
}
`;
```

### How to add to taskpane.js
Find the line where the main R execution string starts (look for `const rCode = \`` or similar near chart dispatch).
Add `${SHARED_STAT_HELPERS_R}` at the very beginning of that template literal, before any chart function definitions.

### Caller pattern after migration

**2-group caller:**
```r
res <- sato_run_stats_2group(
  group1_data = group1_data, group2_data = group2_data,
  group_names = c(groups_with_data[1], groups_with_data[2]),
  x_label = as.character(x_val),
  statistical_test = statistical_test, variance_test = variance_test,
  stat_symbol_type = stat_symbol_type,
  paired = paired,
  error_label = error_label,
  desc_lines = desc_lines
)
stat_text_results <- c(stat_text_results, res$result_text)
sig_label <- res$sig_label
p_val <- res$p_val
```

**3+-group caller:**
```r
res <- sato_run_stats_ngroup(
  anova_data = anova_data,   # data.frame(group, value, [subject_id])
  x_label = as.character(x_val),
  statistical_test = statistical_test, post_hoc_test = post_hoc_test,
  dunnett_control = dunnett_control,
  stat_symbol_type = stat_symbol_type,
  paired = paired,
  error_label = error_label,
  desc_lines = desc_lines
)
stat_text_results <- c(stat_text_results, res$result_text)
omnibus_p <- res$omnibus_p
```

---

## Refactoring Steps (in order)

### Phase 1 — Define helpers (no behavior change)
- [x] Step 1: Add `SHARED_STAT_HELPERS_R` JS constant (code above) to `taskpane.js` — added at line ~22688
- [x] Step 2: Prepend `${SHARED_STAT_HELPERS_R}` at start of `plotCode` template literal (line ~22688)
- [x] Step 3: Build verified — `npm run build` compiled with 0 errors

### Phase 2 — Migrate chart types (one at a time, verify each)
- [ ] Step 4: Migrate `line_grouped_error_raw` — replace 2-group block (~line 15857–16019) and 3+-group block (~line 16022–16338) with helper calls
- [ ] Step 5: Migrate `lq_survival_grouped` — replace 2-group block and 3+-group block with helper calls
- [ ] Step 6: Migrate `box_grouped_dot` — replace inline stat block with helper calls
- [ ] Step 7: Migrate `violin_grouped_dot` — replace inline stat block with helper calls
- [ ] Step 8: Migrate `bar_grouped_error_dot` — replace all 3 copies (v1/v2/v3) with helper calls
- [ ] Step 9: Migrate `box`, `box_dot`, `violin_dot`, `bar_error_dot` — replace inline stat blocks

### Phase 3 — Verify and clean up
- [ ] Step 10: Test all chart types with 2-group data (paired + unpaired)
- [ ] Step 11: Test all chart types with 3+-group data (ANOVA + Kruskal-Wallis paths)
- [ ] Step 12: Verify Dunnett now works in manual mode (was broken before)
- [ ] Step 13: Verify paired support in bar_grouped_error_dot (was missing before)
- [ ] Step 14: Remove old inline stat blocks

### Phase 2 — Migrate Group A chart types (one at a time, verify each)
- [ ] Step 6: Migrate `line_grouped_error_raw` inline stat block → call helpers
- [ ] Step 7: Migrate `lq_survival_grouped` inline stat block → call helpers
- [ ] Step 8: Migrate `box` stat text block → call helpers
- [ ] Step 9: Migrate `box_dot` stat text block → call helpers
- [ ] Step 10: Migrate `violin_dot` stat text block → call helpers
- [ ] Step 11: Migrate `bar_error_dot` stat text block → call helpers
- [ ] Step 12: Migrate `box_grouped_dot` stat text block → call helpers
- [ ] Step 13: Migrate `violin_grouped_dot` stat text block → call helpers
- [ ] Step 14: Migrate `bar_grouped_error_dot` stat text block → call helpers

### Phase 3 — Verify and clean up
- [ ] Step 15: Test all chart types with 2-group data (paired + unpaired)
- [ ] Step 16: Test all chart types with 3+-group data (ANOVA + Kruskal-Wallis paths)
- [ ] Step 17: Test LQ with 2-group paired data
- [ ] Step 18: Test LQ with 3+-group vbracket
- [ ] Step 19: Remove old inline stat blocks

### Known bugs to verify fixed after migration
- [ ] **Dunnett ignored in manual mode** — manual parametric + Dunnett still runs Tukey. Will be fixed because shared helper correctly branches on `post_hoc_test` value from the start.
- [ ] **Paired test missing in bar_grouped_error_dot v1/v2** — will be fixed because shared helper always supports paired.
- [ ] **Variance text wording inconsistency** — "equal" vs "equal variances" — will be standardized in helper.

---

## Key Risks

1. **`get_sig_symbol()` is re-defined inside each chart's R block** — it needs to also be in the shared helpers (or defined once globally in SHARED_STAT_HELPERS_R)
2. **`line_grouped_error_raw` stat results feed into vbracket parsing** — result_text format must stay compatible with the vbracket line-parsing regex
3. **`lq_survival_grouped` stat results feed into vbracket AND into per-dose symbol placement** — the helper must return enough info for both uses
4. **Variable naming** — each chart uses slightly different local variable names (e.g., `x_val` vs `d_val`, `groups_with_data` vs `groups_at_d`) — helpers must use their own internal names
5. **`stat_text_results` global accumulation** — each chart accumulates results across x-axis values via `stat_text_results <- c(stat_text_results, result_text)` — helpers return per-call result_text and caller accumulates

---

## Files to Modify

- `src/taskpane/taskpane.js` — main change (define helpers, replace inline blocks)
- No HTML changes needed

---

## Notes

- `sato_add_statistics_to_plot()` (line ~13168) handles VISUAL brackets using ggpubr — it is a SEPARATE concern from the stat text helpers above. Do NOT merge them. Both will continue to exist after refactoring.
- The educational R code generators (`generateSingleGroupStatisticalCode`, `generateCategoryByComparisonsCode`) are JavaScript functions generating R code strings — they are NOT part of this refactoring (they generate standalone R scripts for users, not in-app execution).
- Priority order: line_grouped_error_raw and lq_survival_grouped first (Phase 2 Steps 6-7), since those are the ones most likely to get new chart types added nearby.
