library(readxl)
library(dplyr)
data<-read_xlsx("Conagra Data 20-24 Geography.xlsx")

# 1. Work on Type of Substitute
#Getting the substitute Summary for top categories
substitute_summary <- table(data$`Type Of Substitute`)

#Getting the top substitutes
top_substitutes <- head(sort(substitute_summary, decreasing = TRUE), 15)
top_substitutes

# Create a new column 'category_sub' with renamed categories
data <- data %>%
  mutate(category_sub = case_when(
    grepl("PLANT", data$`Type Of Substitute`) ~ "PLANT",
    grepl("SOY",data$`Type Of Substitute`) ~ "SOY",
    grepl("VEGGIE",data$`Type Of Substitute`) ~ "VEGGIE",
    grepl("SEITAN",data$`Type Of Substitute`) ~ "SEITAN",
    grepl("JACKFRUIT",data$`Type Of Substitute`) ~ "JACKFRUIT",
    grepl("GRAINS",data$`Type Of Substitute`) ~ "GRAINS",
    grepl("N/A",data$`Type Of Substitute`) ~ "N/A",
    grepl("VALUE",data$`Type Of Substitute`) ~ "N/A",
    TRUE ~ "Others"
  ))
table(data$category_sub)

# Remove rows with "N/A" in the new column
data_sub <- data %>%
  filter(!(category_sub %in% "N/A"))
table(data$category_sub)

# Perform regression analysis
model_sub <- lm(data_sub$`Unit Sales` ~ category_sub + Geography + category_sub*Geography+ data_sub$`Price per Unit`, data = data_sub)
#summary
summary(model_sub)

model_sub <- lm(data_sub$`Dollar Sales` ~ category_sub + category_sub*category_sub, data = data_sub)
#summary
summary(model_sub)


# 2. Work on Top forms 
#top form
form_summary <- table(data$Form)
#Getting the top forms
top_form <- sort(form_summary, decreasing = TRUE)
top_form
formlist<- c("BURGER", "PATTY", "LINK","NUGGET","CRUMBLE","GROUND","STRIP","MEATBALL","SLICE","WING","TENDER")

# Create a new column 'category_sub' with renamed categories
data_form <- data %>%
  mutate(form_category = case_when(
    grepl("BURGER", data$Form) ~ "BURGER",
    grepl("PATTY", data$Form) ~ "PATTY",
    grepl("LINK", data$Form) ~ "LINK",
    grepl("NUGGET", data$Form) ~ "NUGGET",
    grepl("CRUMBLE", data$Form) ~ "CRUMBLE",
    grepl("GROUND", data$Form) ~ "GROUND",
    grepl("STRIP", data$Form) ~ "STRIP",
    grepl("MEATBALL", data$Form) ~ "MEATBALL",
    grepl("SLICE", data$Form) ~ "SLICE",
    grepl("WING", data$Form) ~ "WING",
    grepl("TENDER", data$Form) ~ "TENDER",
    TRUE ~ "Others"
  ))
table(data_form$form_category)

# Perform regression analysis
data_form$form_category <- factor(data_form$form_category)
data_form$Geography <- factor(data_form$Geography)
model_form <- lm(data_form$`Unit Sales` ~ form_category + Geography + form_category*Geography+ data_form$`Price per Unit`, data = data_form)
#summary
summary(model_form)
