library(tidyquant)
library(dplyr)
library(xlsx)
library(tidyr)
library(gt)

rm(list = ls())
gc()

setwd("C:/Users/maran/Documents/Data Projects/Fund Investment Tool")

wb <- loadWorkbook(file = "Investment Analysis Tool.xlsx")


funds <- read.csv("funds.csv", stringsAsFactors = F)
tickers <- funds %>% pull(ticker)



# # Historical Prices - only needs to be run once
# stocks_historical <- tq_get(tickers,
#                             from = "2010-01-01",
#                             to = "2019-12-31",
#                             get = "stock.prices")
# 
# 
# annual_returns <- stocks_historical %>%
#   filter(!is.na(close)) %>%
#   group_by(symbol) %>%
#   tq_transmute(select = close,
#                mutate_fun = periodReturn,
#                period = "yearly",
#                type = "arithmetic")
# 
# avg_annual_returns <- annual_returns %>%
#   group_by(symbol) %>%
#   summarize(avg_return = mean(yearly.returns),
#             obs = n()) %>%
#   data.frame()
# 
# removeSheet(wb, sheetName = "avg annual returns")
# newSheet <- createSheet(wb, sheetName="avg annual returns")
# addDataFrame(avg_annual_returns, newSheet, row.names = F)



# year to date & 6 month rolling returns analysis
# needs to be run regularly

six_months <- today() - months(6)
year_start <- as.Date("2020-01-01")

earliest_date <- min(six_months, year_start)
today <- today()

stocks <- tq_get(tickers,
                 from = earliest_date,
                 to = today,
                 get = "stock.prices")

# Calculate return year to date
ytd_return <- stocks %>%
  group_by(symbol) %>%
  filter(date >= year_start) %>%
  filter(!is.na(close)) %>%
  tq_transmute(select = close,
               mutate_fun = periodReturn,
               period = "yearly",
               type = "arithmetic") %>%
  mutate(type = "ytd") %>%
  rename(date_pulled = date,
         return = yearly.returns,
         ticker = symbol) %>%
  select(ticker, date_pulled, return, type)

# Calculate rolling 6 month return
sixmo_return <- stocks %>%
  group_by(symbol) %>%
  filter(date >= six_months) %>%
  filter(!is.na(close)) %>%
  tq_transmute(select = close,
               mutate_fun = periodReturn,
               period = "yearly",
               type = "arithmetic") %>%
  mutate(type = "rolling_6_mo") %>%
  rename(date_pulled = date,
         return = yearly.returns,
         ticker = symbol) %>%
  select(ticker, date_pulled, return, type)

# compile year to date and 6 month returns, reshape wide
returns <- bind_rows(ytd_return, sixmo_return) %>%
  pivot_wider(names_from = type, values_from = return) %>%
  ungroup() 

# adjust ytd and 6 month returns by expense ratio and merge with fund details
# 'adjusted' refers to return minus expense ratio
percent_year <- as.numeric(today - year_start)/365

returns_full <- full_join(funds, returns, by = "ticker") %>% 
  mutate(ytd_adj       = ytd - expense_ratio * percent_year,
         rolling_adj   = rolling_6_mo - expense_ratio / 2,
         rank_ytd      = rank(desc(ytd_adj)),
         rank_rolling = rank(desc(rolling_adj))) %>%
  arrange(rank_ytd)

# Calculate return summary by fund attributes
returns_class <- returns_full %>%
  group_by(asset_class, size, type, location) %>%
  summarize(avg_ytd = mean(ytd_adj),
            avg_6mo = mean(rolling_6_mo),
            funds = n()) %>%
  arrange(desc(avg_ytd))


# Table for display with index funds
index_tickers <- c("VFIAX", "VTSAX", "VBTLX")

index_rows <- returns_full %>% 
  select(-ytd, -rolling_6_mo, -date_pulled) %>%
  filter(ticker %in% index_tickers) %>%
  mutate(row_type = "Benchmark Funds")

returns_table <- returns_full %>%
  filter(rank_ytd <= 10  | ticker %in% index_tickers) %>%
  select(-ytd, -rolling_6_mo, -date_pulled) %>%
  mutate(row_type = "Top Performing Funds") %>%
  bind_rows(index_rows) %>%
  gt(groupname_col = "row_type") %>%
  fmt_percent(columns = vars(ytd_adj, rolling_adj), decimals = 1) %>%
  fmt_percent(columns = vars(expense_ratio), decimals = 2) 

returns_table



# Add back into Excel file
returns_full <- data.frame(returns_full)
removeSheet(wb, sheetName = "ytd returns")
newSheet <- createSheet(wb, sheetName="ytd returns")
addDataFrame(returns_full, newSheet, row.names = F)

returns_class <- data.frame(returns_class)
removeSheet(wb, sheetName = "class returns")
newSheet <- createSheet(wb, sheetName="class returns")
addDataFrame(returns_class, newSheet, row.names = F)

saveWorkbook(wb, "Investment Analysis Tool.xlsx")
