# Predictive Analytics for Conagra’s Gardein Brand in the U.S. Meat Substitute Market

A comprehensive analysis of market trends, competitive positioning, and predictive modeling for Conagra’s Gardein brand, with recommendations for expanding market share in the U.S. plant-based meat substitute industry.

- [Shaista](https://github.com/AShaista)
- [Yash Tated](https://github.com/Yashtated) 
- Umair Safdar
- Akshat Bhandari
- Jinhyuk Yim
- Yuyu Ndiema
- Nuna Tetteh
  

## Table of Contents

- [Project Overview](#project-overview)
- [About Conagra and Gardein](#about-conagra-and-gardein)
- [Market Landscape](#market-landscape)
  - [Major Competitors](#major-competitors)
  - [Market Size and Growth](#market-size-and-growth)
  - [Geographical Trends](#geographical-trends)
- [Data and Methodology](#data-and-methodology)
- [Strategic Recommendations](#strategic-recommendations)
  - [Expanding Product Portfolio](#expanding-product-portfolio)
  - [Regional Penetration](#regional-penetration)
  - [Pricing Strategy](#pricing-strategy)
- [Appendix: Data Analysis Highlights](#appendix-data-analysis-highlights)


---

## Project Overview

This project applies predictive analytics to understand and enhance the market position of Conagra’s Gardein brand in the U.S. meat substitute sector. Using multi-year sales data, competitive benchmarking, and regression modeling, the analysis identifies actionable strategies for increasing Gardein’s market share and optimizing its product and pricing mix.[1][3][4]

---

## About Conagra and Gardein

**Conagra Brands** is a major U.S. food company with a diverse portfolio, including iconic brands like Healthy Choice, Marie Callender’s, and Chef Boyardee. In 2014, Conagra acquired **Gardein**, a plant-based protein brand founded by Yves Potvin. Gardein offers over 30 vegan and vegetarian products, including meatless alternatives for chicken, beef, pork, and fish, as well as soups and ready meals.[1][5][11]

- Gardein is positioned as a leader in the plant-based segment, with a focus on health-conscious, digitally savvy, and affluent consumers.
- In 2024, Conagra continues to invest in plant-based innovation, expanding Gardein’s offerings and entering new categories such as plant-based soups and breakfast items.[5][6][11]

---

## Market Landscape

### Major Competitors

Gardein faces competition from established and emerging brands:

- **Morningstar Farms** (Kellogg)
- **Beyond Meat**
- **Impossible Foods**
- **Tofurky**
- **Field Roast**
- **Lightlife**
- **Dr. Praeger’s**
- **Boca** (Kraft Heinz)[1][5][7]

### Market Size and Growth

- The U.S. meat substitute market is valued at $8 billion (2023) and is projected to reach $57.6 billion globally by 2033 (CAGR: 14.5%).[1]
- Gardein holds the #4 market share in the U.S., with 18% of meat substitute users reporting usage, 23% brand popularity, and 59% aided brand awareness.[7]
- Conagra’s plant-based segment is expected to grow by 20% in revenue, reaching $1.2 billion in 2024.[6]

### Geographical Trends

- Meat substitute sales are highest in the Northeast, followed by the Southeast and Midsouth regions.
- Gardein’s sales are strong in Midsouth, Plains, South Central, and Southeast, but lag in California, Great Lakes, West, and East regions—areas with significant growth potential.[1]

---

## Data and Methodology

- **Data Sources:** Multi-year retail sales data (2020-2024), product-level and region-level breakdowns, and macroeconomic indicators (inflation, interest rates, COVID-19 impact).[1][3][4]
- **Tools:** R for data cleaning, transformation, and regression modeling (see `Conagara-Project-Final.R` and `Predictive-Analytics-R-Script.R`).
- **Modeling:** Linear regression models were used to assess the impact of product attributes, pricing, regional factors, and distribution on unit sales. Dummy variables captured year and COVID-19 effects.[3][4]

---

## Strategic Recommendations

### Expanding Product Portfolio

- Launch Gardein products in missing categories to capture unmet demand and broaden consumer appeal.
- Focus on high-potential segments, where Gardein already has a foothold but can increase share.[1][5][9]

### Regional Penetration

- Invest in marketing and distribution in underperforming but high-potential regions.
- Tailor products and promotions to regional preferences and meal occasions.[1]

### Pricing Strategy

- Maintain value-based pricing for premium products, with targeted promotions and digital coupons to attract new consumers.[8]
- Monitor price elasticity: If demand is inelastic, modest price increases may boost revenue; if elastic, focus on volume growth through competitive pricing and promotions.[1][8]

---

## Appendix: Data Analysis Highlights

- Regression models identified significant predictors of unit sales: ACV-weighted distribution, price per unit, flavor, region, and macroeconomic factors (CPI, interest rates).
- Dollar sales and unit sales trends were analyzed by product, geography, and time period, revealing seasonal and regional patterns.
- Data cleaning and merging scripts (`Combining-code.R`) ensured robust, consistent datasets for analysis.[2][3][4]




