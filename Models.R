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

# Convert columns to factors
data$`Product Type Assigned`<- factor(data$`Product Type Assigned`)
data$`Assigned Flavor` <- factor(data$`Assigned Flavor`)
data$Geography <- factor(data$Geography)
data$`Assigned Brand` <- factor(data$`Assigned Brand`)
data$Form<- factor(data$Form)
data$Package<- factor(data$Package)
data$Year<-factor(data$Year)

subset_data <- data %>%
  group_by(Geography, `Assigned Flavor`,`Assigned Brand`,`Product Type Assigned`) %>%
  summarise(DollarSales = sum(`Dollar Sales`), UnitSales = sum(`Unit Sales`))

#Models
full <- lm(`Unit Sales` ~ Geography+ `Assigned Brand` +`Assigned Flavor`+`Product Type Assigned`+Form +Package, data=data)
anova(full)
start <- lm(`Unit Sales` ~ Geography, data=data)
null <- lm(`Unit Sales` ~ 1, data=data)
summary(full)

# Backward Elimination using AIC
backward<- step(full, scope=list(lower=null), direction = "backward", k = 2, test="F")

model_subset <- lm(`UnitSales` ~ Geography+`Assigned Brand`*`Assigned Flavor`+`Assigned Flavor` +Geography*`Assigned Flavor`, data=subset_data)
summary(model_subset)

model_full <- lm(`Unit Sales` ~ Geography+`Assigned Brand`*`Assigned Flavor`+`Assigned Flavor` +Geography*`Assigned Flavor`, data=data)
summary(model_full)

model_form <- lm(`Unit Sales` ~ Form + Geography + Form*Geography, data = data)
summary(model_form)

model_flavour <- lm(`Unit Sales` ~ `Assigned Flavor` + Geography + `Assigned Flavor`*Geography , data = data)
summary(model_flavour)

model_acv<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution` , data = data)
summary(model_acv)


data$`Assigned Brand` <- relevel(data$`Assigned Brand`, ref = "GARDEIN") 

model_acv<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ `Assigned Brand`*`ACV Weighted Distribution` , data = data)
summary(model_acv)

model_acv1<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`Assigned Brand`*`ACV Weighted Distribution` , data = data)
summary(model_acv1)

model_acv1<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`Assigned Brand`*`ACV Weighted Distribution` , data = data)
summary(model_acv1)

# Changing the reference category for the brands and product type
data$`Product Type Assigned` <- relevel(data$`Product Type Assigned`, ref = "Meat")
data$`Assigned Brand` <- relevel(data$`Assigned Brand`, ref = "GARDEIN")
model_prod<-lm(`Unit Sales` ~ `Price per Unit` + `Product Type Assigned`*`Assigned Brand` ,data=data)
summary(model_prod)

subset_Data<-data[data$`Assigned Brand`=='GARDEIN',]
subset_Data_1<-data[!data$`Assigned Brand`=='OTHERS',]
subset_Data_1$`Assigned Brand` <- relevel(subset_Data_1$`Assigned Brand`, ref = "GARDEIN")
model_acv1<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution` , data = subset_Data)
summary(model_acv1)

model_acv2<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution` , data = subset_Data_1)
summary(model_acv2)

model_acv3<- lm(`Unit Sales` ~   Geography + `ACV Weighted Distribution`+ Geography*`ACV Weighted Distribution`+CPI+`Interest Rate` , data = data)
summary(model_acv3)

subset2<-subset(subset_Data,Geography %in% c("Northeast", "Southeast"))
subset2<-subset_Data[subset_Data$Geography=="Southeast",]

# Calculate the average price per unit
average_price_per_unit <- mean(subset2$`Price per Unit`, na.rm = TRUE)

# Print the average price per unit
print(average_price_per_unit)
