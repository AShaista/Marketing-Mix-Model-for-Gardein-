library(readxl)
library(dplyr)
#require(mosaic)
#require(knitr)
#require(car)
#require("leaps")
#require("MASS")
#require(olsrr)
options(scipen = 999)

data<-read_xlsx("Conagra Data 20-24 Geography 3.xlsx")

summary(data)

# Select only numeric columns from the dataset
numeric_data <- data[, sapply(data, is.numeric)]

# Compute the correlation matrix
correlation_matrix <- cor(numeric_data)

# Print the correlation matrix
print(correlation_matrix)

# Assuming your correlation matrix is stored in the correlation_matrix variable

# Find the row and column indices of correlations above +0.75 or below -0.75
high_corr_indices <- which(correlation_matrix > 0.75 & correlation_matrix != 1, arr.ind = TRUE)

# Get the column names and corresponding correlations for high correlations
high_corr_values <- data.frame(Column1 = rownames(correlation_matrix)[high_corr_indices[, 1]],
                               Column2 = colnames(correlation_matrix)[high_corr_indices[, 2]],
                               Correlation = correlation_matrix[high_corr_indices])

# Print the high and low correlation values with column names
print(high_corr_values)
