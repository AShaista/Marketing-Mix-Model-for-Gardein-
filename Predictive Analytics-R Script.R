library(readxl)
library(lubridate)
library(tidyverse)
library(dplyr)
library(broom)
library(margins)
library(knitr)
library(pbapply)
library(parallel)
library(coefplot)
library(AER)
library(webshot)
library(openxlsx)


#####################Full Data Set#####################


#Read the Dataset
Conagra_Data_20_24_Complete_1 <- read_excel("C:/Users/Suleman Dawood Ahmad/Desktop/Predictive Project-SA/Conagra Data 20-24 Complete 1.xlsx")
View(Conagra_Data_20_24_Complete_1)

#Check Column Names
names(Conagra_Data_20_24_Complete_1)

#Subset to Keep Relevant Columns
Conagra_Data_20_24_Complete_subset <- subset(Conagra_Data_20_24_Complete_1, select = c("Week Ending", "Geography", "Base Unit Sales", "Price per Unit", "Assigned Flavor", "ACV Weighted Distribution"))

#Rename N_F/S Column
colnames(Conagra_Data_20_24_Complete_subset)[2] <- "Region"
colnames(Conagra_Data_20_24_Complete_subset)[5] <- "Flavor_Scent"
colnames(Conagra_Data_20_24_Complete_subset)[1] <- "Week"

#Create Year Variable
Conagra_Data_20_24_Complete_subset$Year <- year(as.Date(Conagra_Data_20_24_Complete_subset$Week))
Conagra_Data_20_24_Complete_subset <- Conagra_Data_20_24_Complete_subset %>% mutate(covid_indicator=ifelse(Year %in% c(2020, 2021), "Covid", "No Covid"))

#Exclude 2024
Conagra_Data_20_24_Complete_subset <- Conagra_Data_20_24_Complete_subset[Conagra_Data_20_24_Complete_subset$Year!=2024,]

#Create Dummy Variables for Year
Conagra_Data_20_24_Complete_subset <- fastDummies::dummy_cols(Conagra_Data_20_24_Complete_subset, c("Year"))

# Remove spaces from column names
colnames(Conagra_Data_20_24_Complete_subset)
colnames(Conagra_Data_20_24_Complete_subset) <- gsub(" ", "_", colnames(Conagra_Data_20_24_Complete_subset))

#Remove the US Total Region
Conagra_Data_20_24_Complete_subset <- Conagra_Data_20_24_Complete_subset[Conagra_Data_20_24_Complete_subset$Region!="Total US - Multi Outlet + Conv",]

#Convert Flavors to Factor
Conagra_Data_20_24_Complete_subset$Flavor_Scent <- factor(Conagra_Data_20_24_Complete_subset$Flavor_Scent)


#Suppose you want to change the base level to 'Region_A'
Conagra_Data_20_24_Complete_subset$Region <- relevel(factor(Conagra_Data_20_24_Complete_subset$Region), ref = "California - Standard - Multi Outlet + Conv")
Conagra_Data_20_24_Complete_subset$Flavor_Scent <- relevel(Conagra_Data_20_24_Complete_subset$Flavor_Scent, ref = "Regular")

#Merge Macro Data with Conagra Subset
Conagra_Data_20_24_Complete_subset_mr <- right_join(macro_week, Conagra_Data_20_24_Complete_subset)

#QA Check the Rows
nrow(Conagra_Data_20_24_Complete_subset)
nrow(Conagra_Data_20_24_Complete_subset_mr)

#Regression with Time Trend
flavor_reg <- lm(Base_Unit_Sales~ ACV_Weighted_Distribution+Flavor_Scent*Region+CPI+Price_per_Unit+`Interest Rate`+Year_2021+Year_2022+Year_2020,
                 Conagra_Data_20_24_Complete_subset_mr)

#Summary of the Regression 
(summ_flavor_reg <- summary(flavor_reg))

#Convert to Tidy Output
view(flavor_reg$coefficients)


tidy_summary <- broom::tidy(flavor_reg)


#Number of Significant Variables
count(tidy_summary[tidy_summary$p.value<0.05,])


# Create a data frame for the regression summary
summary_df <- data.frame(
  Variable = tidy_summary$term,
  Coefficient = tidy_summary$estimate,
  Std.Error = tidy_summary$std.error,
  t.value = tidy_summary$statistic,
  Pr = tidy_summary$p.value
)

# Add R-squared and Adjusted R-squared to the summary data frame
summary_df <- rbind(summary_df, data.frame(Variable = "R-squared", Coefficient = summ_flavor_reg$r.squared, Std.Error = "", t.value = "", Pr = ""))
summary_df <- rbind(summary_df, data.frame(Variable = "Adjusted R-squared", Coefficient = summ_flavor_reg$adj.r.squared, Std.Error = "", t.value = "", Pr = ""))
summary_df <- rbind(summary_df, data.frame(Variable = "No. of Obs", Coefficient = summ_flavor_reg$df[2], Std.Error = "", t.value = "", Pr = ""))


# Convert relevant columns to numeric
summary_df$Coefficient <- as.numeric(summary_df$Coefficient)
summary_df$Std.Error <- as.numeric(summary_df$Std.Error)
summary_df$t.value <- as.numeric(summary_df$t.value)
summary_df$Pr <- as.numeric(summary_df$Pr)


#Round off the Values for 2 Decimal Points
summary_df[, -1] <- round(summary_df[, -1], 2)

# Convert p-values to stars for significance
summary_df$Pr <- ifelse(summary_df$Pr < 0.05, "***", ifelse(summary_df$Pr < 0.1, "**", ifelse(summary_df$Pr < 0.15, "*", "")))


# Generate HTML table and save it to a file
generate_regression_html <- function(summary_df, output_file, title) {
  # Convert the summary data frame to HTML using kable with title and caption
  html_table <- knitr::kable(summary_df, "html", digits = 3, align = "c")
  
  # Create the full HTML content with title and table
  full_html <- paste0("<html><head><title>", title, "</title></head><body><h1>", title, "</h1>", html_table, "</body></html>")
  
  # Write the HTML content to a file
  writeLines(full_html, output_file)
  
  # Open the HTML file in a web browser
  browseURL(output_file)
  
}


# Call the function with your summary data frame and desired output file name
generate_regression_html(summary_df, "regression_summary_fulldata.html", title="Flavor Scent Effect across Region-Complete Data")


#Get Average Marginal Return of the Variables
model_vars <- all.vars(formula(flavor_reg))

# Filter out the response variable
model_vars <- model_vars[model_vars != "Base_Unit_Sales"]

# Initialize an empty list to store marginal effects
marginal_effects <- list()

# Loop through each variable in the model
for (var in model_vars) {
  # Calculate marginal effects using margins function
  marginal <- margins(flavor_reg, variables = var)
  
  # Store the marginal effect in the list
  marginal_effects[[var]] <- summary(marginal)
}

# Print the list of marginal effects
print(marginal_effects)

marginal_effects_df <- do.call(rbind, marginal_effects)
view(marginal_effects_df)


##################Impossible Brand###############################

#Check Column Names
names(Conagra_Data_20_24_Complete_1)

#Subset to Keep Relevant Columns
Conagra_Data_20_24_Complete_subset_imp <- subset(Conagra_Data_20_24_Complete_1, select = c("Week Ending", "Geography", "Base Unit Sales", "Price per Unit", "Assigned Flavor", "ACV Weighted Distribution", "Brand"))

#Filter for Impossible Brand
view(unique(Conagra_Data_20_24_Complete_subset_imp$Brand))

Conagra_Data_20_24_Complete_subset_imp <- Conagra_Data_20_24_Complete_subset_imp[Conagra_Data_20_24_Complete_subset_imp$Brand=="IMPOSSIBLE",]

#Rename N_F/S Column
colnames(Conagra_Data_20_24_Complete_subset_imp)[2] <- "Region"
colnames(Conagra_Data_20_24_Complete_subset_imp)[5] <- "Flavor_Scent"
colnames(Conagra_Data_20_24_Complete_subset_imp)[1] <- "Week"
Conagra_Data_20_24_Complete_subset_imp$Brand <- NULL

#Create Year Variable
Conagra_Data_20_24_Complete_subset_imp$Year <- year(as.Date(Conagra_Data_20_24_Complete_subset_imp$Week))
Conagra_Data_20_24_Complete_subset_imp <- Conagra_Data_20_24_Complete_subset_imp %>% mutate(covid_indicator=ifelse(Year %in% c(2020, 2021), "Covid", "No Covid"))

#Exclude 2024
Conagra_Data_20_24_Complete_subset_imp <- Conagra_Data_20_24_Complete_subset_imp[Conagra_Data_20_24_Complete_subset_imp$Year!=2024,]

#Create Dummy Variables for Year
Conagra_Data_20_24_Complete_subset_imp <- fastDummies::dummy_cols(Conagra_Data_20_24_Complete_subset_imp, c("Year"))

# Remove spaces from column names
colnames(Conagra_Data_20_24_Complete_subset_imp)
colnames(Conagra_Data_20_24_Complete_subset_imp) <- gsub(" ", "_", colnames(Conagra_Data_20_24_Complete_subset_imp))

#Remove the US Total Region
Conagra_Data_20_24_Complete_subset_imp <- Conagra_Data_20_24_Complete_subset_imp[Conagra_Data_20_24_Complete_subset_imp$Region!="Total US - Multi Outlet + Conv",]

#Convert Flavors to Factor
Conagra_Data_20_24_Complete_subset_imp$Flavor_Scent <- factor(Conagra_Data_20_24_Complete_subset_imp$Flavor_Scent)


#Suppose you want to change the base level to 'Region_A'
Conagra_Data_20_24_Complete_subset_imp$Region <- relevel(factor(Conagra_Data_20_24_Complete_subset_imp$Region), ref = "California - Standard - Multi Outlet + Conv")
Conagra_Data_20_24_Complete_subset_imp$Flavor_Scent <- relevel(Conagra_Data_20_24_Complete_subset_imp$Flavor_Scent, ref = "Regular")

#Merge Macro Data with Conagra Subset
Conagra_Data_20_24_Complete_subset_imp_mr <- right_join(macro_week, Conagra_Data_20_24_Complete_subset_imp)

#QA Check the Rows
nrow(Conagra_Data_20_24_Complete_subset_imp)
nrow(Conagra_Data_20_24_Complete_subset_imp_mr)

#Regression with Time Trend
flavor_reg_imp <- lm(Base_Unit_Sales~ ACV_Weighted_Distribution+Flavor_Scent*Region+CPI+Price_per_Unit+`Interest Rate`+Year_2021+Year_2022+Year_2020,
                 Conagra_Data_20_24_Complete_subset_imp_mr)

#Summary of the Regression 
(summ_flavor_reg_imp <- summary(flavor_reg_imp))

#Convert to Tidy Output
view(flavor_reg_imp$coefficients)


tidy_summary_imp <- broom::tidy(flavor_reg_imp)


#Number of Significant Variables
count(tidy_summary_imp[tidy_summary_imp$p.value<0.05,])


# Create a data frame for the regression summary
summary_df_imp <- data.frame(
  Variable = tidy_summary_imp$term,
  Coefficient = tidy_summary_imp$estimate,
  Std.Error = tidy_summary_imp$std.error,
  t.value = tidy_summary_imp$statistic,
  Pr = tidy_summary_imp$p.value
)

# Add R-squared and Adjusted R-squared to the summary data frame
summary_df_imp <- rbind(summary_df_imp, data.frame(Variable = "R-squared", Coefficient = summ_flavor_reg_imp$r.squared, Std.Error = "", t.value = "", Pr = ""))
summary_df_imp <- rbind(summary_df_imp, data.frame(Variable = "Adjusted R-squared", Coefficient = summ_flavor_reg_imp$adj.r.squared, Std.Error = "", t.value = "", Pr = ""))
summary_df_imp <- rbind(summary_df_imp, data.frame(Variable = "No. of Obs", Coefficient = summ_flavor_reg_imp$df[2], Std.Error = "", t.value = "", Pr = ""))


# Convert relevant columns to numeric
summary_df_imp$Coefficient <- as.numeric(summary_df_imp$Coefficient)
summary_df_imp$Std.Error <- as.numeric(summary_df_imp$Std.Error)
summary_df_imp$t.value <- as.numeric(summary_df_imp$t.value)
summary_df_imp$Pr <- as.numeric(summary_df_imp$Pr)


#Round off the Values for 2 Decimal Points
summary_df_imp[, -1] <- round(summary_df_imp[, -1], 2)

# Convert p-values to stars for significance
summary_df_imp$Pr <- ifelse(summary_df_imp$Pr < 0.05, "***", ifelse(summary_df_imp$Pr < 0.1, "**", ifelse(summary_df_imp$Pr < 0.15, "*", "")))


# Call the function with your summary data frame and desired output file name
generate_regression_html(summary_df_imp, "regression_summary_impossible.html", title="Flavor Scent Effect across Region-IMPOSSIBLE Brand")


#Get Average Marginal Return of the Variables
model_vars_imp <- all.vars(formula(flavor_reg_imp))

# Filter out the response variable
model_vars_imp <- model_vars_imp[model_vars_imp != "Base_Unit_Sales"]

# Initialize an empty list to store marginal effects
marginal_effects_imp <- list()

# Loop through each variable in the model
for (var in model_vars_imp) {
  # Calculate marginal effects using margins function
  marginal <- margins(flavor_reg_imp, variables = var)
  
  # Store the marginal effect in the list
  marginal_effects_imp[[var]] <- summary(marginal)
}

# Print the list of marginal effects
print(marginal_effects_imp)

marginal_effects_df_imp <- do.call(rbind, marginal_effects_imp)
view(marginal_effects_df_imp)


##################Gardien Brand###############################

#Check Column Names
names(Conagra_Data_20_24_Complete_1)

#Subset to Keep Relevant Columns
Conagra_Data_20_24_Complete_subset_gdn <- subset(Conagra_Data_20_24_Complete_1, select = c("Week Ending", "Geography", "Base Unit Sales", "Price per Unit", "Assigned Flavor", "ACV Weighted Distribution", "Brand"))

#Filter for Impossible Brand
view(unique(Conagra_Data_20_24_Complete_subset_gdn$Brand))

Conagra_Data_20_24_Complete_subset_gdn <- Conagra_Data_20_24_Complete_subset_gdn[Conagra_Data_20_24_Complete_subset_gdn$Brand=="GARDEIN",]

#Rename N_F/S Column
colnames(Conagra_Data_20_24_Complete_subset_gdn)[2] <- "Region"
colnames(Conagra_Data_20_24_Complete_subset_gdn)[5] <- "Flavor_Scent"
colnames(Conagra_Data_20_24_Complete_subset_gdn)[1] <- "Week"
Conagra_Data_20_24_Complete_subset_gdn$Brand <- NULL

#Create Year Variable
Conagra_Data_20_24_Complete_subset_gdn$Year <- year(as.Date(Conagra_Data_20_24_Complete_subset_gdn$Week))
Conagra_Data_20_24_Complete_subset_gdn <- Conagra_Data_20_24_Complete_subset_gdn %>% mutate(covid_indicator=ifelse(Year %in% c(2020, 2021), "Covid", "No Covid"))

#Exclude 2024
Conagra_Data_20_24_Complete_subset_gdn <- Conagra_Data_20_24_Complete_subset_gdn[Conagra_Data_20_24_Complete_subset_gdn$Year!=2024,]

#Create Dummy Variables for Year
Conagra_Data_20_24_Complete_subset_gdn <- fastDummies::dummy_cols(Conagra_Data_20_24_Complete_subset_gdn, c("Year"))

# Remove spaces from column names
colnames(Conagra_Data_20_24_Complete_subset_gdn)
colnames(Conagra_Data_20_24_Complete_subset_gdn) <- gsub(" ", "_", colnames(Conagra_Data_20_24_Complete_subset_gdn))

#Remove the US Total Region
Conagra_Data_20_24_Complete_subset_gdn <- Conagra_Data_20_24_Complete_subset_gdn[Conagra_Data_20_24_Complete_subset_gdn$Region!="Total US - Multi Outlet + Conv",]

#Convert Flavors to Factor
Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent <- factor(Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent)


#Suppose you want to change the base level to 'Region_A'
Conagra_Data_20_24_Complete_subset_gdn$Region <- relevel(factor(Conagra_Data_20_24_Complete_subset_gdn$Region), ref = "California - Standard - Multi Outlet + Conv")
Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent <- relevel(Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent, ref = "Regular")

#Merge Macro Data with Conagra Subset
Conagra_Data_20_24_Complete_subset_gdn_mr <- right_join(macro_week, Conagra_Data_20_24_Complete_subset_gdn)

#QA Check the Rows
nrow(Conagra_Data_20_24_Complete_subset_gdn)
nrow(Conagra_Data_20_24_Complete_subset_gdn_mr)

#Regression with Time Trend
flavor_reg_gdn <- lm(Base_Unit_Sales~ ACV_Weighted_Distribution+Flavor_Scent*Region+CPI+Price_per_Unit+`Interest Rate`+Year_2021+Year_2022+Year_2020,
                     Conagra_Data_20_24_Complete_subset_gdn_mr)

#Summary of the Regression 
(summ_flavor_reg_gdn <- summary(flavor_reg_gdn))

#Convert to Tidy Output
view(flavor_reg_gdn$coefficients)


tidy_summary_gdn <- broom::tidy(flavor_reg_gdn)


#Number of Significant Variables
count(tidy_summary_imp[tidy_summary_imp$p.value<0.05,])


# Create a data frame for the regression summary
summary_df_gdn <- data.frame(
  Variable = tidy_summary_gdn$term,
  Coefficient = tidy_summary_gdn$estimate,
  Std.Error = tidy_summary_gdn$std.error,
  t.value = tidy_summary_gdn$statistic,
  Pr = tidy_summary_gdn$p.value
)

# Add R-squared and Adjusted R-squared to the summary data frame
summary_df_gdn <- rbind(summary_df_gdn, data.frame(Variable = "R-squared", Coefficient = summ_flavor_reg_gdn$r.squared, Std.Error = "", t.value = "", Pr = ""))
summary_df_gdn <- rbind(summary_df_gdn, data.frame(Variable = "Adjusted R-squared", Coefficient = summ_flavor_reg_gdn$adj.r.squared, Std.Error = "", t.value = "", Pr = ""))
summary_df_gdn <- rbind(summary_df_gdn, data.frame(Variable = "No. of Obs", Coefficient = summ_flavor_reg_gdn$df[2], Std.Error = "", t.value = "", Pr = ""))


# Convert relevant columns to numeric
summary_df_gdn$Coefficient <- as.numeric(summary_df_gdn$Coefficient)
summary_df_gdn$Std.Error <- as.numeric(summary_df_gdn$Std.Error)
summary_df_gdn$t.value <- as.numeric(summary_df_gdn$t.value)
summary_df_gdn$Pr <- as.numeric(summary_df_gdn$Pr)


#Round off the Values for 2 Decimal Points
summary_df_gdn[, -1] <- round(summary_df_gdn[, -1], 2)

# Convert p-values to stars for significance
summary_df_gdn$Pr <- ifelse(summary_df_gdn$Pr < 0.05, "***", ifelse(summary_df_gdn$Pr < 0.1, "**", ifelse(summary_df_gdn$Pr < 0.15, "*", "")))


# Call the function with your summary data frame and desired output file name
generate_regression_html(summary_df_gdn, "regression_summary_gardein.html", title="Flavor Scent Effect across Region-GARDEIN Brand")


#Get Average Marginal Return of the Variables
model_vars_gdn <- all.vars(formula(flavor_reg_gdn))

# Filter out the response variable
model_vars_gdn <- model_vars_gdn[model_vars_gdn != "Base_Unit_Sales"]

# Initialize an empty list to store marginal effects
marginal_effects_gdn <- list()

# Loop through each variable in the model
for (var in model_vars_gdn) {
  # Calculate marginal effects using margins function
  marginal <- margins(flavor_reg_gdn, variables = var)
  
  # Store the marginal effect in the list
  marginal_effects_gdn[[var]] <- summary(marginal)
}

# Print the list of marginal effects
print(marginal_effects_gdn)

marginal_effects_df_gdn <- do.call(rbind, marginal_effects_gdn)
view(marginal_effects_df_gdn)


#Output Marginal Effects of All the Models
write.xlsx(list(marginal_effects_df, marginal_effects_df_gdn, marginal_effects_df_imp), sheetName=list("Complete Data", "GARDEIN Brand", "IMPOSSIBLE Brand"), 
           "C:/Users/Suleman Dawood Ahmad/Desktop/Predictive Project-SA/Outputs/Marginal Effects for All three Models.xlsx")

# Save the current R working environment
save.image(file = "C:/Users/Suleman Dawood Ahmad/Desktop/Predictive Project-SA/R Script and R Objects/Conagra Models-R Objects.RData")
