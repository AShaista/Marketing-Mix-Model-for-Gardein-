library(readxl)
library(ggplot2)
options(scipen=999)
#read data
FZ_Poultry_2020 <- read_excel("Fz_Rfg Poultry_POS_2020.xlsx")
FZ_Poultry_2021 <- read_excel("Fz_Rfg Poultry_POS_2021.xlsx")
FZ_Poultry_2022<- read_excel("Fz_Rfg Poultry_POS_2022.xlsx")
FZ_Poultry_2023 <- read_excel("Fz_Rfg Poultry_POS_2023.xlsx")
FZ_Poultry_2024 <- read_excel("Fz_Rfg Poultry_POS_2024.xlsx")

#combine data
combined_data <- rbind(FZ_Poultry_2020,FZ_Poultry_2021,FZ_Poultry_2022,FZ_Poultry_2023)
category_to_omit <- "Total US - Multi Outlet + Conv"
combined_data <- combined_data[!combined_data$Geography == category_to_omit,]
combined_data$Date<- format(combined_data$Date, "%M-%D%-%Y")
combined_data$Year <- as.integer(format(combined_data$Date, "%Y"))

#------------Aggregate Dollar sales by year-----------
dollar_sales_by_year <- aggregate(`Dollar Sales` ~ Year, data = combined_data, sum)

# Plot the graph with y-axis in Millions
ggplot(dollar_sales_by_year, aes(x = Year, y = `Dollar Sales` / 1e6)) +
  geom_line() +
  geom_point() +
  labs(title = "Dollar Sales of Poultry (2020-2023)",
       x = "Year",
       y = "Dollar Sales (Millions)")

#------------Aggregate Volume sales by year-----------
Volume_sales_by_year <- aggregate(`Volume Sales` ~ Year, data = combined_data, sum)

# Plot the graph with y-axis in thousands
ggplot(Volume_sales_by_year, aes(x = Year, y = `Volume Sales` / 1e6)) +
  geom_line() +
  geom_point() +
  labs(title = "Volume Sales of Poultry (2020-2023)",
       x = "Year",
       y = "Volume Sales (Millions)")

#Processed Poultry
FZ_Processed_Poultry_2020 <- read_excel("Fz_Rfg Processed Poultry_POS_2020.xlsx")
FZ_Processed_Poultry_2021 <- read_excel("Fz_Rfg Processed Poultry_POS_2021.xlsx")
FZ_Processed_Poultry_2022 <- read_excel("Fz_Rfg Processed Poultry_POS_2022.xlsx")
FZ_Processed_Poultry_2023 <- read_excel("Fz_Rfg Processed Poultry_POS_2023.xlsx")
FZ_Processed_Poultry_2024 <- read_excel("Fz_Rfg Processed Poultry_POS_2024.xlsx")

#----------------------Aggregate Processed poultry----------------------
FZ_Processed_Poultry_2023$Date<- format(as.Date(FZ_Processed_Poultry_2023$Date,"%m-%d-%y"),"%m-%d-%y")
Combined_Data <- rbind(FZ_Processed_Poultry_2020,FZ_Processed_Poultry_2021,FZ_Processed_Poultry_2022,FZ_Processed_Poultry_2023)
Combined_Data <- Combined_Data[!Combined_Data$Geography == category_to_omit,]
Combined_Data$Year <- as.integer(format(Combined_Data$Date, "%Y"))

#------------Aggregate Dollar sales by year-----------
dollar_sales_by_year <- aggregate(`Dollar Sales` ~ Year, data = Combined_Data, sum)
FZ23<-FZ_Processed_Poultry_2023[!FZ_Processed_Poultry_2023$Geography==category_to_omit,]
dollar_sales_23 <- sum(FZ23$`Dollar Sales`)
  #aggregate(`Dollar Sales`, data = FZ_Processed_Poultry_2023, sum)

# Plot the graph with y-axis in thousands
ggplot(dollar_sales_by_year, aes(x = Year, y = `Dollar Sales` / 1e6)) +
  geom_line() +
  geom_point() +
  labs(title = "Dollar Sales of Poultry (2020-2023)",
       x = "Year",
       y = "Dollar Sales (Millions)")

#------------Aggregate Volume sales by year-----------
Volume_sales_by_year <- aggregate(`Volume Sales` ~ Year, data = combined_data, sum)

# Plot the graph with y-axis in thousands
ggplot(Volume_sales_by_year, aes(x = Year, y = `Volume Sales` / 1000)) +
  geom_line() +
  geom_point() +
  labs(title = "Volume Sales of Poultry (2020-2023)",
       x = "Year",
       y = "Volume Sales (Thousands)") 





subset_data <- subset(FZ_Processed_Poultry_2023, Geography != 'Total US - Multi Outlet + Conv')

# Calculate the total sales
total_dollar_sales1 <- sum(subset_data$`Volume Sales`)
total_dollar_sales1

# Create a data frame
sales_data <- data.frame(
  Year = c(2020, 2021, 2022, 2023),
  Sales = c(981352331, 1023222461,  1023255957,1092578975)
)

ggplot(sales_data, aes(x = Year, y = Sales / 1e6)) +
  geom_line() +
  geom_point() +
  labs(title = "Volume Sales of Processed Poultry (2020-2023)",
       x = "Year",
       y = "Volume Sales (Millions)") 
