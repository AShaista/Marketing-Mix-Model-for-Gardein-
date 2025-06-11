library(readxl)
library(ggplot2)
options(scipen=999)

FZ_Sub_Meat_2020 <- read_excel("Fz_Rfg Substitute Meat_POS_2020.xlsx")
head(FZ_Sub_Meat_2020)
FZ_Sub_Meat_2020$Tofurky<- ifelse(grepl("^TOFURKY", FZ_Sub_Meat_2020$Product), 1, 0)
FZ_Sub_Meat_2020$Tofurky
subset_tofurky<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$Tofurky==1,]

sales_by_geography <- aggregate(subset_tofurky$`Dollar Sales`, by = list(Geography = subset_tofurky$Geography), FUN = sum)
category_to_omit <- "Total US - Multi Outlet + Conv"
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Tofurky Dollar sales Vs Geography-2020")

#fIELD ROAST
FZ_Sub_Meat_2020$field<- ifelse(grepl("^FIELD", FZ_Sub_Meat_2020$Product), 1, 0)
subset_field<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$field==1,]

sales_by_geography <- aggregate(subset_field$`Dollar Sales`, by = list(Geography = subset_field$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Field Roast Dollar sales Vs Geography-2020")

# BOCA 
FZ_Sub_Meat_2020$boca<- ifelse(grepl("^BOCA", FZ_Sub_Meat_2020$Product), 1, 0)
subset_boca<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$boca==1,]

sales_by_geography <- aggregate(subset_boca$`Dollar Sales`, by = list(Geography = subset_boca$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "BOCA Dollar sales Vs Geography-2020")

#DR PRAEGERS
FZ_Sub_Meat_2020$dr<- ifelse(grepl("^DR PRAEGERS", FZ_Sub_Meat_2020$Product), 1, 0)
subset_dr<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$dr==1,]

sales_by_geography <- aggregate(subset_dr$`Dollar Sales`, by = list(Geography = subset_dr$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "DR Praegers Dollar sales Vs Geography-2020")

#Beyond meat
FZ_Sub_Meat_2020$beyond<- ifelse(grepl("^BEYOND", FZ_Sub_Meat_2020$Product), 1, 0)
subset_beyond<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$beyond==1,]

sales_by_geography <- aggregate(subset_beyond$`Dollar Sales`, by = list(Geography = subset_beyond$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Beyond meat Dollar sales Vs Geography-2020")

#GARDEIN
FZ_Sub_Meat_2020$gardein<- ifelse(grepl("^GARDEIN", FZ_Sub_Meat_2020$Product), 1, 0)
subset_gardein<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$gardein==1,]

sales_by_geography <- aggregate(subset_gardein$`Dollar Sales`, by = list(Geography = subset_gardein$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Gardein Dollar sales Vs Geography-2020")

#Morningstar
FZ_Sub_Meat_2020$morning<- ifelse(grepl("^MORNINGSTAR", FZ_Sub_Meat_2020$Product), 1, 0)
subset_moring<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$morning==1,]

sales_by_geography <- aggregate(subset_moring$`Dollar Sales`, by = list(Geography = subset_moring$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Mornignstar Dollar sales Vs Geography-2020")

#impossible
FZ_Sub_Meat_2020$imp<- ifelse(grepl("^IMPOSSIBLE", FZ_Sub_Meat_2020$Product), 1, 0)
subset_imp<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$imp==1,]

sales_by_geography <- aggregate(subset_imp$`Dollar Sales`, by = list(Geography = subset_imp$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Impossible Dollar sales Vs Geography-2020")

#private label
FZ_Sub_Meat_2020$pvt<- ifelse(grepl("^PRIVATE", FZ_Sub_Meat_2020$Product), 1, 0)
subset_pvt<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$pvt==1,]

sales_by_geography <- aggregate(subset_pvt$`Dollar Sales`, by = list(Geography = subset_pvt$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Private Label sales Vs Geography-2020")

#lightlife
FZ_Sub_Meat_2020$light<- ifelse(grepl("^LIGHTLIFE", FZ_Sub_Meat_2020$Product), 1, 0)
subset_light<-FZ_Sub_Meat_2020[FZ_Sub_Meat_2020$light==1,]

sales_by_geography <- aggregate(subset_light$`Dollar Sales`, by = list(Geography = subset_light$Geography), FUN = sum)
sales_by_geography <- sales_by_geography[!sales_by_geography$Geography == category_to_omit,]

#bar chart
ggplot(sales_by_geography, aes(x = Geography, y = x, fill = Geography)) +
  geom_bar(stat = "identity") +
  labs(title = "Lightlife sales Vs Geography-2020")
