# Created:  20230720
# Author:   Gage Rowden

#########################################

# List of required packages
required_packages <- c("tidyverse", "readxl", "writexl", "openxlsx")

# Function to check if a package is installed
is_package_installed <- function(package_name) {
  return(requireNamespace(package_name, quietly = TRUE))
}

# Check if each package is installed and install it if needed
for (package_name in required_packages) {
  if (!is_package_installed(package_name)) {
    message(paste0("Installing ", package_name, "..."))
    install.packages(package_name)
  }
}
rm(required_packages, package_name)

# Prompt the user to set the working directory.
working_directory <- readline("Please set the working directory: ")
setwd(working_directory)
rm(working_directory)
  
# Prompt the user to identify the input file.
input_file <- ""
while (input_file == "") {
  input_file <- readline("Enter the input file name (WITHOUT extension): ")
  input_file <- paste0(input_file, ".xlsx")
  if (!file.exists(input_file)) {
    cat("File not found. Please try again. \n")
    input_file <- ""
  }
}

# Define the number of rows of metadata the user wants to remove.
num_rows <- ""
while (num_rows == "") {
  num_rows <- as.integer(readline("Enter the number or rows you'd like to remove: "))
  if (is.na(num_rows)) {
    cat("Invalid input.")
    num_rows <- ""
  }
}

# Read the Excel file into R.
data <- read_excel(input_file, sheet = 2)

# Remove metadata.
tidy_data <- data[-(1:num_rows-1), -1] %>%
  na.omit(tidy_data)
rm(data)

# Set the first row as column names.
col_names <- tidy_data[1, ]
tidy_data <- tidy_data[-1, ]
colnames(tidy_data) <- col_names
rm(col_names)

# Add leading "0" before single digits in column names.
colnames(tidy_data) <- gsub(" X(\\d)$", " X0\\1", colnames(tidy_data))

# Identify and handle duplicate column names.
dup_cols <- colnames(tidy_data)[duplicated(colnames(tidy_data))]
if (length(dup_cols) > 0) {
  
  # Add suffix to duplicate column names
  for (col in dup_cols) {
    indices <- which(colnames(tidy_data) == col)
    colnames(tidy_data)[indices] <- paste0(col, "_", indices)
  }
}
rm(col, dup_cols, indices)

# Rename the first column as "Time"
tidy_data <- tidy_data %>%
  rename(Time = 1)

# Rearrange columns to group replicates of the same sample
tidy_data <- tidy_data %>%
  select(Time, order(colnames(tidy_data), decreasing = FALSE))

# Remove suffixes from column names
colnames(tidy_data) <- gsub("_\\d+$", "", colnames(tidy_data))

# Designate the integers used to calculate how the data will be cut
cycles <- length(unique(tidy_data$Time))    # Number of cycles
num_rows <- cycles                          # This will change after sending one data type to a data frame
reads <- length(which(tidy_data$Time==0))   # Number of types of data (e.g. Raw, Normalized, or Derivative)

# Create a data frame with only the "Time" column with no duplicates
time_df <- data.frame(unique(tidy_data$Time)) %>%
  rename(Time = 1)

# Create separate data frames for different read types
i = 1
while (i <= reads) {
  if (num_rows == cycles) {
    df <- cbind(time_df, tidy_data[(num_rows - cycles):num_rows, -1])
    assign(paste0("df", i), df)
    num_rows <- num_rows + cycles
  } else {
    df <- cbind(time_df, tidy_data[(1 + num_rows - cycles):num_rows, -1])
    assign(paste0("df", i), df)
    num_rows <- num_rows + cycles
  }
  i <- i + 1
}
rm(tidy_data, time_df, cycles, num_rows)

# Export the organized data
existing_file <- openxlsx::loadWorkbook(input_file)

# Function to write a data frame to a new sheet in the workbook
write_to_sheet <- function(df, sheet_name, existing_file) {
  openxlsx::addWorksheet(existing_file, sheetName = sheet_name)
  openxlsx::writeData(existing_file, sheet = sheet_name, df, startRow = 1, startCol = 1, rowNames = FALSE)
}

# Write each data frame to a new sheet in the workbook
i = 1
while (i <= reads) {
  df_name <- paste0("df", i)
  df <- get(df_name)
  sheet_name <- paste0("Data", i)
  write_to_sheet(df, sheet_name, existing_file)
  i = i + 1
}
rm(df, df_name, i, reads, sheet_name)

# Save the modified file
openxlsx::saveWorkbook(existing_file, input_file, overwrite = TRUE)

# Open the file for the user to view
if (Sys.info()["sysname"] == "Windows") {
  shell.exec(input_file)
} else if (Sys.info()["sysname"] == "Darwin") {
  system(paste("open", input_file))
}
