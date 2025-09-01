# -----------------------------
# 0. Load necessary libraries
# -----------------------------
# Make sure you have these packages installed:
# install.packages(c("readxl", "survey", "dplyr", "gt"))

library(readxl)
library(survey)
library(dplyr)
library(gt)

# -----------------------------
# 1. Load Data
# -----------------------------
# The file path needs to be correct for your system.
# The original script will not be runnable without access to this file.
df <- read_excel("E:\\1_Research\\COC\\coc1.xlsx")


df$v005 <- df$v005 / 1000000 # scale weights

# -----------------------------
# 2. Create survey design
# -----------------------------
design <- svydesign(
  ids = ~psuid,
  strata = ~strataid,
  weights = ~v005,
  data = df,
  nest = TRUE
)

# -----------------------------
# 3. Variables to analyze
# -----------------------------
# Using only the variables present in the sample dataframe
vars_to_analyze <- c(
  "coc", "anc4", "sba", "pnc", "autonomy",
  "resp_agegrp", "hus_agegrp", "hus_occgrp",
  "media_exposure_cat", "pregnancy_order",
  "multiple_birth", "age_1st_birth", "birth_interval_cat",
  "desire_children", "v106", "v714", "v190", "v024", "v025", "v467d"
)

# -----------------------------
# 4. Compute weighted frequencies & proportions
# -----------------------------
results <- list()

for (var in vars_to_analyze) {
  if (!var %in% colnames(df)) next # skip if variable missing
  
  cat("Processing:", var, "\n")
  
  form <- as.formula(paste0("~", var))
  
  # survey-weighted table
  freq_table <- svytable(form, design, na.rm = TRUE)
  
  df_freq <- as.data.frame(freq_table)
  colnames(df_freq) <- c("Category", "Weighted_Freq")
  
  df_freq$Category <- as.character(df_freq$Category)
  
  # FIX: Calculate proportion (0 to 1) instead of percentage.
  # The gt `fmt_percent` function will handle the multiplication by 100.
  df_freq$Proportion <- df_freq$Weighted_Freq / sum(df_freq$Weighted_Freq)
  
  df_freq$Variable <- var
  
  results[[var]] <- df_freq
}

# Combine all results into a single dataframe
combined <- do.call(rbind, results) %>%
  arrange(Variable, Category)

# -----------------------------
# 5. Create GT table
# -----------------------------
gt_table <- combined %>%
  select(Variable, Category, Weighted_Freq, Proportion) %>%
  group_by(Variable) %>%
  gt(rowname_col = "Category") %>%
  tab_header(
    title = "Univariate Distribution of Key Variables (Survey-Weighted)",
    subtitle = "Bangladesh DHS-style Analysis"
  ) %>%
  tab_options(
    table.font.size = 10,
    column_labels.font.size = 10,
    heading.title.font.size = 12,
    heading.subtitle.font.size = 10,
    row_group.font.size = 12
  ) %>%
  # FIX: Updated `vars()` to `c()` to remove deprecation warning.
  # Changed decimals to 0 for frequency counts.
  fmt_number(columns = c(Weighted_Freq), decimals = 0) %>%
  # FIX: `fmt_percent` now correctly scales the Proportion column.
  fmt_percent(columns = c(Proportion), decimals = 1) %>%
  cols_align(align = "center") %>%
  cols_label(
    Weighted_Freq = "Weighted Frequency",
    Proportion = "Percentage (%)" # Label the proportion column
  ) %>%
  tab_style(
    style = list(cell_text(weight = "bold")),
    locations = cells_column_labels()
  ) %>%
  tab_style(
    style = list(cell_fill(color = "#f8f8f8")),
    locations = cells_row_groups()
  )

# -----------------------------
# 6. Display GT table
# -----------------------------
gt_table






# -----------------------------
# 0. Load necessary libraries
# -----------------------------
# Make sure you have these packages installed:
# install.packages(c("readxl", "survey", "dplyr", "gt", "tidyr"))

library(readxl)
library(survey)
library(dplyr)
library(gt)
library(tidyr)


# -----------------------------
# 1. Load Data
# -----------------------------
# The file path needs to be correct for your system.
# The original script will not be runnable without access to this file.
df <- read_excel("E:\\1_Research\\COC\\coc1.xlsx")

# Ensure the outcome variable is a factor for analysis
df$coc <- as.factor(df$coc)
# FIX: Create a numeric 0/1 version of the outcome for svymean
# This prevents dimension mismatch errors with confint()
positive_level <- levels(df$coc)[2] # Assume the second level is the "Yes" case
df$coc_numeric <- as.numeric(df$coc == positive_level)

colnames(df) <- tolower(colnames(df)) # lowercase column names
df$v005 <- df$v005 / 1000000 # scale weights

# -----------------------------
# 2. Create survey design
# -----------------------------
design <- svydesign(
  ids = ~psuid,
  strata = ~strataid,
  weights = ~v005,
  data = df,
  nest = TRUE
)

# -----------------------------
# 3. Define outcome and predictor variables
# -----------------------------
outcome_var <- "coc"
outcome_var_numeric <- "coc_numeric" # Use the numeric version for calculations
predictor_vars <- c(
  "anc4", "sba", "pnc", "autonomy",
  "resp_agegrp", "hus_agegrp", "hus_occgrp",
  "media_exposure_cat", "pregnancy_order",
  "multiple_birth", "age_1st_birth", "birth_interval_cat",
  "desire_children", "v106", "v714", "v190", "v024", "v025", "v467d"
)

# -----------------------------
# 4. Compute bivariate statistics
# -----------------------------
results_list <- list()

for (var in predictor_vars) {
  if (!var %in% colnames(df) || all(is.na(df[[var]]))) next
  
  cat("Processing:", var, "\n")
  
  # A. Chi-squared test (uses the original factor variable)
  chisq_formula <- as.formula(paste0("~", outcome_var, " + ", var))
  chisq_test <- svychisq(chisq_formula, design)
  p_value <- chisq_test$p.value
  
  # B. Row percentages and confidence intervals
  by_formula <- as.formula(paste0("~", var))
  # FIX: Use the numeric outcome variable in the formula
  mean_formula <- as.formula(paste0("~", outcome_var_numeric))
  
  # Calculate proportions. This will now return one mean and one SE.
  row_props <- svyby(mean_formula, by = by_formula, design = design, FUN = svymean, na.rm = TRUE)
  
  # Calculate confidence intervals. This will now have the correct dimensions.
  ci <- confint(row_props)
  
  # Combine results into a clean data frame
  res_df <- as.data.frame(row_props)
  
  # FIX: Rename columns based on the new structure and add CI columns
  colnames(res_df) <- c("Category", "coc_Yes_Prop", "SE")
  res_df$coc_Yes_CI_Lower <- ci[,1]
  res_df$coc_Yes_CI_Upper <- ci[,2]
  
  # Calculate "No" proportion and CI
  res_df$coc_No_Prop <- 1 - res_df$coc_Yes_Prop
  res_df$coc_No_CI_Lower <- 1 - res_df$coc_Yes_CI_Upper
  res_df$coc_No_CI_Upper <- 1 - res_df$coc_Yes_CI_Lower
  
  res_df$Variable <- var
  res_df$p_value <- p_value
  
  results_list[[var]] <- res_df
}

# Combine all results
combined <- do.call(rbind, results_list)

# -----------------------------
# 5. Prepare data for GT table
# -----------------------------
gt_data <- combined %>%
  # Format percentages and CIs into single strings
  mutate(
    coc_No = paste0(format(round(coc_No_Prop * 100, 1), nsmall = 1), "% (", 
                    format(round(coc_No_CI_Lower * 100, 1), nsmall = 1), " - ", 
                    format(round(coc_No_CI_Upper * 100, 1), nsmall = 1), ")"),
    coc_Yes = paste0(format(round(coc_Yes_Prop * 100, 1), nsmall = 1), "% (", 
                     format(round(coc_Yes_CI_Lower * 100, 1), nsmall = 1), " - ", 
                     format(round(coc_Yes_CI_Upper * 100, 1), nsmall = 1), ")"),
    p_value_formatted = if_else(p_value < 0.001, "<0.001", as.character(round(p_value, 3)))
  ) %>%
  select(Variable, Category, coc_No, coc_Yes, p_value_formatted)

# -----------------------------
# 6. Create and Display GT table
# -----------------------------
gt_table <- gt_data %>%
  group_by(Variable) %>%
  gt(rowname_col = "Category") %>%
  tab_header(
    title = "Bivariate Analysis of Predictors for Continuity of Care (coc)",
    subtitle = "Row Percentages with 95% Confidence Intervals"
  ) %>%
  tab_spanner(
    label = "Continuity of Care (coc)",
    columns = c(coc_No, coc_Yes)
  ) %>%
  cols_label(
    coc_No = "No",
    coc_Yes = "Yes",
    p_value_formatted = "P-value"
  ) %>%
  cols_align(align = "center") %>%
  tab_style(
    style = list(cell_text(weight = "bold")),
    locations = cells_column_labels(columns = everything())
  ) %>%
  tab_style(
    style = list(cell_fill(color = "#f9f9f9")),
    locations = cells_row_groups()
  ) %>%
  opt_table_font(font = "Arial")

gt_table


