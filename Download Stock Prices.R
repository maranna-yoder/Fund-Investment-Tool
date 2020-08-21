library(tidyquant)
library(dplyr)
library(xlsx)
library(tidyr)
library(gt)

rm(list = ls())
gc()

setwd("C:/Users/maran/Documents/Data Projects/Fund Investment Tool")

wb <- loadWorkbook(file = "./public/Vanguard Top Performers.xlsx")


funds <- read.csv("./public/funds.csv", stringsAsFactors = F)
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




### year to date, 3 month rolling, and 1 month rolling returns analysis
### needs to be run regularly


# date info
three_months <- today() - months(3)
one_month    <- today() - months(1)
two_weeks    <- today() - days(14)
year_start  <- as.Date("2020-01-01")

earliest_date <- min(three_months, year_start)
today <- today()


# pull stock data
stocks <- tq_get(tickers,
                 from = earliest_date,
                 to = today,
                 get = "stock.prices")


# Function to calculate return from today to date
FundReturnFunction <- function(stock_data, start_date, type_col){
  
  period_return <- stock_data %>%
    group_by(symbol) %>%
    filter(date >= start_date) %>%
    filter(!is.na(close)) %>%
    tq_transmute(select = close,
                 mutate_fun = periodReturn,
                 period = "yearly",
                 type = "arithmetic") %>%
    mutate(type = type_col) %>%
    rename(date_pulled = date,
           return = yearly.returns,
           ticker = symbol) %>%
    select(ticker, date_pulled, return, type)

  return(period_return)
  
}

ytd_return        <- FundReturnFunction(stocks, year_start, "ytd")
rolling3_return   <- FundReturnFunction(stocks, three_months, "rolling_3_mo")
rolling1_return   <- FundReturnFunction(stocks, one_month, "rolling_1_mo")
rolling2wk_return <- FundReturnFunction(stocks, two_weeks, "rolling_2_wk")


# compile calculated returns, reshape wide
returns <- bind_rows(ytd_return, rolling3_return, rolling1_return, rolling2wk_return) %>%
  pivot_wider(names_from = type, values_from = return) %>%
  ungroup() 


# adjust ytd and rolling returns by expense ratio and merge with fund details
# 'adjusted' refers to return minus prorated expense ratio
# the true expense ratio calculation involves converting to an annualized basis
# rather than pro-rating, but I omit it here as an approximation.

percent_year <- as.numeric(today - year_start)/365

returns_full <- full_join(funds, returns, by = "ticker") %>% 
  mutate(ytd_adj       = ytd - expense_ratio * percent_year,
         rolling3_adj   = rolling_3_mo - expense_ratio / 4,
         rolling1_adj   = rolling_1_mo - expense_ratio / 12,
         rolling2wk_adj = rolling_2_wk - expense_ratio / 26,
         rank_ytd      = rank(desc(ytd_adj), ties.method = "first"),
         rank_rolling3 = rank(desc(rolling3_adj), ties.method = "first"),
         rank_rolling1 = rank(desc(rolling1_adj), ties.method = "first"),
         rank_rolling2wk = rank(desc(rolling2wk_adj), ties.method = "first")) %>%
  arrange(rank_rolling3)

# Calculate return summary by fund attributes
returns_class <- returns_full %>%
  group_by(asset_class, size, type, location) %>%
  summarize(avg_ytd = mean(ytd_adj),
            avg_3mo = mean(rolling3_adj),
            avg_1mo = mean(rolling1_adj),
            avg_2wk = mean(rolling2wk_adj),
            funds = n()) %>%
  arrange(-avg_1mo) %>%
  ungroup() %>%
  mutate(rank1mo = rank(desc(avg_1mo), ties.method = "first"))


# Table for display with index funds
index_tickers <- c("VFIAX", "VTSAX", "VBTLX")

index_rows <- returns_full %>% 
  select(-ytd, -rolling_3_mo, -rolling_1_mo, -date_pulled) %>%
  filter(ticker %in% index_tickers) %>%
  mutate(row_type = "Benchmark Funds")

returns_table <- returns_full %>%
  filter(rank_rolling3 <= 10  | ticker %in% index_tickers) %>%
  select(-ytd, -rolling_3_mo, -rolling_1_mo, -date_pulled) %>%
  mutate(row_type = "Top Performing Funds") %>%
  bind_rows(index_rows) %>%
  gt(groupname_col = "row_type") %>%
  fmt_percent(columns = vars(ytd_adj, rolling3_adj, rolling1_adj), decimals = 1) %>%
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

saveWorkbook(wb, "./public/Vanguard Top Performers.xlsx")
