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


library(stringr)

combined$Category <- str_to_title(combined$Category)

var_labels <- c(
  coc = "Completed continuum of care (ANC4 + SBA + PNC)",
  anc4 = "At least 4 ANC visits with skilled provider",
  sba = "Skilled Birth Attendant at delivery",
  pnc = "Maternal PNC within 2 days (skilled)",
  autonomy = "Autonomy in health care decision",
  resp_agegrp = "Respondent's Age (years, grouped)",
  hus_agegrp = "Husband/Partner Age (years, grouped)",
  hus_occgrp = "Husband/Partner Occupation (grouped)",
  media_exposure_cat = "Media Exposure",
  pregnancy_order = "Pregnancy order category",
  multiple_birth = "Child is single or multiple birth",
  age_1st_birth = "Age at first birth",
  birth_interval_cat = "Birth interval in months",
  desire_children = "Desire for more children",
  v106 = "Highest educational level",
  v714 = "Respondent currently working",
  v190 = "Wealth index combined",
  v024 = "Division",
  v025 = "Type of place of residence",
  v467d = "Getting medical help: distance to health facility"
)


combined$Variable <- var_labels[combined$Variable]




# -----------------------------
# 5. Create GT table
# -----------------------------
gt_table <- combined %>%
  select(Variable, Category, Weighted_Freq, Proportion) %>%
  rename(
    N = Weighted_Freq,
    `%` = Proportion
  ) %>%
  gt(
    rowname_col = "Category",
    groupname_col = "Variable"  # Makes Variable the row group header
  ) %>%
  tab_header(
    title = "Univariate Distribution of Key Variables (Survey-Weighted)",
    subtitle = "Bangladesh DHS-style Analysis"
  ) %>%
  tab_options(
    table.font.size = 10,
    column_labels.font.size = 10,
    heading.title.font.size = 12,
    heading.subtitle.font.size = 10,
    row_group.font.size = 11,
    row_group.background.color = "#f2f2f2"
  ) %>%
  fmt_number(columns = N, decimals = 0) %>%
  fmt_percent(columns = `%`, decimals = 1) %>%
  cols_align(align = "center") %>%
  cols_label(
    Category = "Category",
    N = "N",
    `%` = "%"
  ) %>%
  tab_style(
    style = cell_text(weight = "bold"),
    locations = cells_column_labels()
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
   "autonomy",
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

library(stringr)

combined$Category <- str_to_title(combined$Category)

# ... [previous code remains the same until the GT table preparation] ...

# Create a mapping of variable names to their labels
var_label_mapping <- c(
  "anc4" = "At least 4 ANC visits with skilled provider",
  "sba" = "Skilled Birth Attendant at delivery",
  "pnc" = "Maternal PNC within 2 days (skilled)",
  "autonomy" = "Autonomy in health care decision",
  "resp_agegrp" = "Respondent's Age (years, grouped)",
  "hus_agegrp" = "Husband/Partner Age (years, grouped)",
  "hus_occgrp" = "Husband/Partner Occupation (grouped)",
  "media_exposure_cat" = "Media Exposure",
  "pregnancy_order" = "Pregnancy order category",
  "multiple_birth" = "Child is single or multiple birth",
  "age_1st_birth" = "Age at first birth",
  "birth_interval_cat" = "Birth interval in months",
  "desire_children" = "Desire for more children",
  "v106" = "Highest educational level",
  "v714" = "Respondent currently working",
  "v190" = "Wealth index combined",
  "v024" = "Division",
  "v025" = "Type of place of residence",
  "v467d" = "Distance to health facility"
)

# Create a mapping for category labels
category_label_mapping <- list(
  "0" = "No",
  "1" = "Yes",
  "Not a Big Problem" = "Not a big problem",
  "Child is Single or Multiple Birth" = "Child is single or multiple birth"
  # Add more category mappings as needed based on your data
)

# 5. Prepare data for GT table
gt_data <- combined %>%
  # Map variable names to labels
  mutate(Variable_Label = ifelse(Variable %in% names(var_label_mapping), 
                                 var_label_mapping[Variable], 
                                 Variable)) %>%
  # Map category values to labels
  mutate(Category_Label = ifelse(as.character(Category) %in% names(category_label_mapping),
                                 category_label_mapping[as.character(Category)],
                                 as.character(Category))) %>%
  # Format percentages and CIs
  mutate(
    coc_No = paste0(format(round(coc_No_Prop * 100, 1), nsmall = 1), "% (", 
                    format(round(coc_No_CI_Lower * 100, 1), nsmall = 1), " - ", 
                    format(round(coc_No_CI_Upper * 100, 1), nsmall = 1), ")"),
    coc_Yes = paste0(format(round(coc_Yes_Prop * 100, 1), nsmall = 1), "% (", 
                     format(round(coc_Yes_CI_Lower * 100, 1), nsmall = 1), " - ", 
                     format(round(coc_Yes_CI_Upper * 100, 1), nsmall = 1), ")"),
    p_value_formatted = if_else(p_value < 0.001, "<0.001", 
                                format(round(p_value, 3), nsmall = 3))
  ) %>%
  select(Variable_Label, Category_Label, coc_No, coc_Yes, p_value_formatted) %>%
  group_by(Variable_Label) %>%
  mutate(group_p_value = first(p_value_formatted)) %>%
  ungroup()

# 6. Create and Display GT table
gt_table <- gt_data %>%
  gt(groupname_col = "Variable_Label") %>%
  tab_header(
    title = "Bivariate Analysis of Predictors for Continuity of Care (coc)",
    subtitle = "Row Percentages with 95% Confidence Intervals"
  ) %>%
  tab_spanner(
    label = "Continuity of Care (coc)",
    columns = c(coc_No, coc_Yes)
  ) %>%
  cols_label(
    Category_Label = "Category",
    coc_No = "No",
    coc_Yes = "Yes",
    group_p_value = "p-value"
  ) %>%
  cols_hide(columns = p_value_formatted) %>%
  cols_align(align = "center") %>%
  tab_style(
    style = list(cell_text(weight = "bold")),
    locations = cells_column_labels()
  ) %>%
  tab_style(
    style = list(cell_fill(color = "#f9f9f9"),
                 cell_text(weight = "bold")),
    locations = cells_row_groups()
  ) %>%
  tab_options(
    row_group.as_column = TRUE,
    data_row.padding = px(2)
  ) %>%
  tab_style(
    style = list(cell_borders(sides = "bottom", color = "grey", weight = px(2))),
    locations = cells_row_groups()
  )

gt_table
