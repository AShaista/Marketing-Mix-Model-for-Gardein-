# Script for combined sales of years, based on brands
library(openxlsx) 
install.packages("lubridate")
library(lubridate)

Combined_Data <- read_excel("Sub Meat 20-23 V2.xlsx")
FZ_Sub_Meat_2024 <- read_excel("Fz_Rfg Substitute Meat_POS_2024.xlsx")

# Assuming df is your data frame and week_ending is the column with the date strings
FZ_Sub_Meat_2024$`Week Ending` <- sub("week ending ", "", FZ_Sub_Meat_2024$`Week Ending`)  # Remove "week ending"

# Convert the column to the date format using lubridate
FZ_Sub_Meat_2024$`Week Ending` <- mdy(FZ_Sub_Meat_2024$`Week Ending`)

# Format the dates to "DD-MM-YYYY" format
FZ_Sub_Meat_2024$`Week Ending` <- format(FZ_Sub_Meat_2024$`Week Ending`, "%d-%m-%Y")

Data <- rbind(Combined_Data,FZ_Sub_Meat_2024)
write.xlsx(Data, file = "Combined 20-24.xlsx", rowNames = FALSE)

write.xlsx(FZ_Sub_Meat_2024, file = "Sub24.xlsx", rowNames = FALSE)
