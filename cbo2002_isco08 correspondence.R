# --- Step 1: Setup ---
# Load the necessary libraries
library(haven)
library(readxl)
library(dplyr)

# --- Step 2: Define File Paths ---
cbo2002_isco88 <- "S:/Tradutores/CBO/cbo2002_ciuo88_2025_macro.xlsm"
isco08_isco88 <- "S:/Tradutores/ISCO/isco08_isco88_alternate_original.xlsx"
isco08_titles <- "S:/Tradutores/ISCO/isco08_titles.sas7bdat"

# --- Step 3: Load All Data Files ---
# CBO to ISCO-88 correspondence
cbo2002_isco88 <- read_excel(cbo2002_isco88, sheet = "cbo2002_isco88")

# ISCO-88 to ISCO-08 bridge file (the link between the two standards)
isco08_isco88 <- read_excel(isco08_isco88, sheet = "isco08_isco88_alternate")

# ISCO-08 titles lookup table
isco08_titles <- read_sas(isco08_titles)

# --- Step 4: Process the Correspondence in Two Parts ---

# Part A: Handle CBOs with an 'alternate_title' (using the 'suggested' ISCO-88)
# We filter for rows where alternate_title is not missing.
cbo_suggested_path <- cbo2002_isco88 %>%
  filter(!is.na(alternate_title)) %>%
  # Join using BOTH the suggested ISCO-88 code and the alternate title
  left_join(
    isco08_isco88,
    by = c("isco88_suggested" = "isco88", "alternate_title" = "alternate_title")
  )

# Part B: Handle CBOs WITHOUT an 'alternate_title' (using the 'oficial' ISCO-88)
# We filter for rows where alternate_title IS missing.
cbo_official_path <- cbo2002_isco88 %>%
  filter(is.na(alternate_title)) %>%
  # Join using ONLY the official ISCO-88 code
  left_join(
    # We only need the ISCO88-ISCO08 link here, no alternate titles needed
    isco08_isco88,
    by = c("isco88_suggested" = "isco88")
  )

# --- Step 5: Combine the two parts and add ISCO-08 titles ---
# Stack the results from Part A and Part B back into one table
cbo2002_isco08 <- bind_rows(cbo_suggested_path, cbo_official_path) %>%
  # Now that we have the isco08 codes, join to get their titles
  left_join(isco08_titles, by = "isco08") %>%
  # This keeps one copy of any row that is identical across all columns.
  distinct() %>%
  # Keep only the final, required columns in the correct order
  select(
    cbo2002, cbo2002_desc, 
    isco88_oficial, isco88_desc_oficial, 
    isco88_suggested, isco88_desc_suggested, 
    isco08, isco08_title, 
    alternate_title
  )

# --- Step 6: Count distinct ISCO-08 codes per CBO-2002 ---
# This creates a summary table to answer your final question
isco08_counts_per_cbo <- cbo2002_isco08 %>%
  group_by(cbo2002, cbo2002_desc) %>%
  summarise(
    distinct_isco08_count = n_distinct(isco08),
    .groups = 'drop' # Recommended practice to ungroup after summarising
  )

# --- Step 7: Display Results ---
cat("--- Correspondence creation complete! ---\n\n")
cat("--- First 10 rows of the final 'cbo2002_isco08' file: ---\n")
print(head(cbo2002_isco08, 10))

cat("\n\n--- Count of distinct ISCO-08 codes for the first 10 CBOs: ---\n")
print(head(isco08_counts_per_cbo, 10))

# --- Exportando para Excel ---
# --- Passo 1: Definir o caminho e o nome do arquivo de destino ---
caminho_destino <- "S:/Tradutores/CBO/isco08_counts_per_cbo.csv"

# --- Passo 2: Exportar o dataframe ---
write_csv(isco08_counts_per_cbo, caminho_destino)
