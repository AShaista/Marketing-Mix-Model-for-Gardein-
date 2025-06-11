library(readxl) 
library(dplyr)  
library(tidyr)  
library(ggplot2) 
library(reshape2)

data <- read_excel("Conagra Data 20-24 Geography 2.xlsx")

data$`Product Type Assigned` <- factor(data$`Product Type Assigned`)
data$Geography <- factor(data$Geography)


heatmap_data <- data %>%
  select('Dollar Sales', Geography, `Product Type Assigned`)

heatmap_melted <- melt(heatmap_data, id.vars = c("Geography", "Product Type Assigned"))

heatmap_plot <- ggplot(data = heatmap_melted, aes(x = Geography, y = `Product Type Assigned`, fill = value,color=value)) +
  geom_tile() +
  scale_fill_gradient(low = "lightblue", high = "red") + scale_color_gradient(low = "lightblue", high = "red") +
  labs(title = "Dollar Sales Heatmap: Geography vs. Type", x = "Geography", y = "Type") +
  theme_minimal()

# Display the heatmap plot
print(heatmap_plot)
