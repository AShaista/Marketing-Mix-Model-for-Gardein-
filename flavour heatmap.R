# Load the required libraries
library(readxl)  # For reading Excel files
library(dplyr)   # For data manipulation
library(tidyr)   # For reshaping data
library(ggplot2) # For plotting
library(reshape2)

# Load your dataset from an Excel file
data <- read_excel("Conagra Data 20-24 Geography 2.xlsx")

# Convert Flavor / Scent and Geography to factors
data$`Product Type Assigned`<- factor(data$`Product Type Assigned`)
data$`Assigned Flavor` <- factor(data$`Assigned Flavor`)
data$Geography <- factor(data$Geography)

subset_data <- data %>%
  group_by(Geography, `Assigned Flavor`) %>%
  summarise(DollarSales = sum(`Dollar Sales`))

summary(subset_data)
# Create a subset of the data for Dollar Sales, Geography, and Flavor
#heatmap_data <- data %>%
 # select('Dollar Sales', Geography, `Assigned Flavor`)


# Melt the data into long format suitable for plotting the heatmap
#heatmap_melted <- melt(subset_data, id.vars = c("Geography",`Assigned Flavor`))

# Create the heatmap plot
heatmap_plot <- ggplot(data = subset_data, aes(x = Geography, y = `Assigned Flavor`, fill = DollarSales,color=DollarSales)) +
  geom_tile() +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +  # Adjust the fill color scale as needed
  scale_color_gradient(low = "lightblue", high = "darkblue") + # Adjust the color scale as needed
  labs(title = "Dollar Sales Heatmap: Geography vs. Flavour", x = "Geography", y = "Flavour") +
  theme_minimal() +
  theme(
    axis.text = element_text(size = 12),  # Adjust the font size for axis labels (x and y)
    axis.title = element_text(size = 14),  # Adjust the font size for axis titles
    plot.title = element_text(size = 16, face = "bold")  # Adjust the font size and style for the plot title
  )

# Display the heatmap plot
print(heatmap_plot)


#prod type
subset_data1 <- data %>%
  group_by(Geography, `Product Type Assigned`) %>%
  summarise(DollarSales = sum(`Dollar Sales`))

summary(subset_data1)

# Create the heatmap plot
heatmap_plot1 <- ggplot(data = subset_data1, aes(x = Geography, y = `Product Type Assigned`, fill = DollarSales,color=DollarSales)) +
  geom_tile() +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +  # Adjust the fill color scale as needed
  scale_color_gradient(low = "lightblue", high = "darkblue") + # Adjust the color scale as needed
  labs(title = "Dollar Sales Heatmap: Geography vs. Product Type", x = "Geography", y = "Product Type") +
  theme_minimal()+
  theme(
    axis.text = element_text(size = 12),  # Adjust the font size for axis labels (x and y)
    axis.title = element_text(size = 14),  # Adjust the font size for axis titles
    plot.title = element_text(size = 16, face = "bold")  # Adjust the font size and style for the plot title
  )

# Display the heatmap plot
print(heatmap_plot1)


subset_data1 <- data %>%
  group_by(Geography, `Assigned Flavor`) %>%
  summarise(UnitSales = sum(`Unit Sales`))

# Create the heatmap plot
heatmap_plot2 <- ggplot(data = subset_data1, aes(x = Geography, y = `Assigned Flavor`, fill = UnitSales,color=UnitSales)) +
  geom_tile() +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +  # Adjust the fill color scale as needed
  scale_color_gradient(low = "lightblue", high = "darkblue") + # Adjust the color scale as needed
  labs(title = "Unit Sales Heatmap: Geography vs. Flavour", x = "Geography", y = "Flavour") +
  theme_minimal()

# Display the heatmap plot
print(heatmap_plot2)
