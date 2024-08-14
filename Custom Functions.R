#Custom Functions for Conagra R Project for Market Growth Analysis

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

#Function for Getting Marginal Effects of the Variables; Outputs Data frame of Average Marginal Effects
marg_effects <- function(input_reg_eq, dependent_var="Base_Unit_Sales") {

model_vars <- all.vars(formula(input_reg_eq))

# Filter out the response variable
model_vars <- model_vars[model_vars != dependent_var]

# Initialize an empty list to store marginal effects
marginal_effects <- list()

# Loop through each variable in the model
for (var in model_vars) {
  # Calculate marginal effects using margins function
  marginal <- margins(input_reg_eq, variables = var)
  
  # Store the marginal effect in the list
  marginal_effects[[var]] <- summary(marginal)
}

# Print the list of marginal effects
print(marginal_effects)

marginal_effects_df <- do.call(rbind, marginal_effects)
view(marginal_effects_df) 

}

