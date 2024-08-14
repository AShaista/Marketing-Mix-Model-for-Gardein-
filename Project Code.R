#Bringing in Required Libraries for Analysis

if (!require("pacman")) install.packages("pacman"); library(pacman)


p_load(readxl, lubridate, tidyverse, dplyr, broom, margins, knitr, pbapply, parallel, coefplot, AER, webshot, openxlsx)

source("D:/Active Work/UTD MSBA/Semesters/Fall 22/Predicitive Analytics for Data Science/Predictive Project-SA/R Script and R Objects/Custom Functions.R")

#####################Full Data Set#####################


#Read the Dataset
Conagra_Data_20_24_Complete <- read_excel("Conagra Dataset.xlsx")


#Check Column Names
names(Conagra_Data_20_24_Complete)

#Subset to Keep Relevant Columns
Conagra_Data_20_24_Complete_subset <- subset(Conagra_Data_20_24_Complete, select = c("Week", "Geography", "Base Unit Sales", "Price per Unit", "Assigned Flavor", "ACV Weighted Distribution", "CCI", "CPI", "Interest Rate","Unemployment Rate"))

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
Conagra_Data_20_24_Complete_subset <- Conagra_Data_20_24_Complete_subset[Conagra_Data_20_24_Complete_subset$Region!="Total",]

#Convert Flavors to Factor
Conagra_Data_20_24_Complete_subset$Flavor_Scent <- factor(Conagra_Data_20_24_Complete_subset$Flavor_Scent)


#Suppose you want to change the base level to 'Region_A'
Conagra_Data_20_24_Complete_subset$Region <- relevel(factor(Conagra_Data_20_24_Complete_subset$Region), ref = "California")
Conagra_Data_20_24_Complete_subset$Flavor_Scent <- relevel(Conagra_Data_20_24_Complete_subset$Flavor_Scent, ref = "Regular")

#Initial Regression Equation
flavor_reg_1 <- lm(Base_Unit_Sales~ Flavor_Scent*Region+Price_per_Unit,
                   Conagra_Data_20_24_Complete_subset)
summary(flavor_reg_1)


#Regression with Time Trend
flavor_reg <- lm(Base_Unit_Sales~ ACV_Weighted_Distribution+Flavor_Scent*Region+CPI+Price_per_Unit+Interest_Rate+Year_2021+Year_2022+Year_2020,
                 Conagra_Data_20_24_Complete_subset)

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

# Call the function with your summary data frame and desired output file name
generate_regression_html(summary_df, "regression_summary_fulldata.html", title="Flavor Scent Effect across Region-Complete Data")


#Get Average Marginal Return of the Variables
marg_effects(flavor_reg)

##################Gardein Brand###############################

#Check Column Names
names(Conagra_Data_20_24_Complete)

#Subset to Keep Relevant Columns
Conagra_Data_20_24_Complete_subset_gdn <- subset(Conagra_Data_20_24_Complete, select = c("Week", "Geography", "Base Unit Sales", "Price per Unit", "Assigned Flavor", "ACV Weighted Distribution", "CCI", "CPI", "Interest Rate", "Unemployment Rate", "Brand"))

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
#Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent <- factor(Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent)


#Suppose you want to change the base level to 'Region_A'
Conagra_Data_20_24_Complete_subset_gdn$Region <- relevel(factor(Conagra_Data_20_24_Complete_subset_gdn$Region), ref = "California")
Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent <- relevel(Conagra_Data_20_24_Complete_subset_gdn$Flavor_Scent, ref = "Regular")


#Regression with Time Trend
flavor_reg_gdn <- lm(Base_Unit_Sales~ ACV_Weighted_Distribution+Flavor_Scent*Region+CPI+Price_per_Unit+Interest_Rate+Year_2021+Year_2022+Year_2020,
                     Conagra_Data_20_24_Complete_subset_gdn)

#Summary of the Regression 
(summ_flavor_reg_gdn <- summary(flavor_reg_gdn))

#Convert to Tidy Output
view(flavor_reg_gdn$coefficients)


tidy_summary_gdn <- broom::tidy(flavor_reg_gdn)


#Number of Significant Variables
count(tidy_summary_gdn[tidy_summary_gdn$p.value<0.05,])


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
marg_effects(flavor_reg_gdn)

#creating a subset to perfrom further analysis
names(Conagra_Data_20_24_Complete_1)
names(Conagra_Data_20_24_Complete_1)[1] <- "Region"
data <- Conagra_Data_20_24_Complete_1[Conagra_Data_20_24_Complete_1$Region!="Total",]

# Convert columns to factors
data$`Product Type Assigned`<- factor(data$`Product Type Assigned`)
data$`Assigned Flavor` <- factor(data$`Assigned Flavor`)
data$Geography <- factor(data$Region)
data$`Assigned Brand` <- factor(data$`Assigned Brand`)
data$Year<-factor(data$Year)


#xreating a subset of data for only gardein
subset_Data<-data[data$`Assigned Brand`=='GARDEIN',]

# creating a subset of data to keep only th top 10 brands 
subset_Data_1<-data[!data$`Assigned Brand`=='OTHERS',]

# creating a reference category for brands
data$`Assigned Brand` <- relevel(data$`Assigned Brand`, ref = "GARDEIN")
subset_Data_1$`Assigned Brand` <- relevel(subset_Data_1$`Assigned Brand`, ref = "GARDEIN")


#Model for ACV comparison of brands and geographies(overall)
model_acv1<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution` , data = data)
summary(model_acv1)

#model for unit sales for only gardein 
model_acv12<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution` , data = subset_Data)
summary(model_acv12)

#model for unit sales for all brands
model_acv3<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution` , data = subset_Data_1)
summary(model_acv3)

#model for unit sales using macro economics parameters.
model_acv4<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution`+CPI+`Interest Rate` , data = data)
summary(model_acv4)



# looking into product pricing vs geography

#creating a subset with only gardein 
data21 <- data[(data$Brand == "GARDEIN" ),]

#creating a subset from plains
data23 <- data21[(data21$Geography == "Plains" ),]

#creating a subset for all other brands
data22 <- data[!(data$Brand == "GARDEIN" ),]

#Models
model_price <- lm(`Unit Sales` ~ Geography  + `Price per Unit` + `Price per Unit` * Geography, data = data)
summary(model_price)

model_price2 <- lm(`Unit Sales`~ Geography + `Price per Unit`+ `Price per Unit` * Geography, data = data21)
summary(model_price2)

model_price3 <- lm(`Unit Sales` ~ Geography + `Price per Unit` + `Price per Unit` * Geography, data = data22)
summary(model_price3)

model_price4 <- lm(`Unit Sales` ~ `Price per Unit`, data = data23)
summary(model_price4)


# looking into the product type and the pricing 

# Changing the reference category for the brands and product type
data$`Product Type Assigned` <- relevel(data$`Product Type Assigned`, ref = "Meat")
data$`Assigned Brand` <- relevel(data$`Assigned Brand`, ref = "GARDEIN")
data$Year<-relevel(data$Year,ref=2020)

#Creating a subset of data for only gardein
subset_Data<-data[data$`Assigned Brand`=='GARDEIN',]

# creating a subset of data to keep only th top 10 brands 
subset_Data_1<-data[!data$`Assigned Brand`=='OTHERS',]

model_prod1<-lm(`Unit Sales` ~ `Price per Unit` +`ACV Weighted Distribution`+ `Product Type Assigned`*Geography,data=subset_Data)
summary(model_prod1)

model_prod2<-lm(`Unit Sales` ~ `Price per Unit` +`ACV Weighted Distribution`+ `Product Type Assigned`*Geography ,data=subset_Data_1)
summary(model_prod2)

model_prod3<-lm(`Unit Sales` ~ `Price per Unit` +`ACV Weighted Distribution`+ `Product Type Assigned`*`Price per Unit` ,data=data)
summary(model_prod3)