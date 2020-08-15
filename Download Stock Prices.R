library(tidyquant)
library(dplyr)
library(xlsx)
library(tidyr)

rm(list = ls())
gc()

setwd("C:/Users/maran/Documents/Data Stuff/Fund Investment Tool")

wb <- loadWorkbook(file = "Investment Analysis Tool.xlsx")


funds <- read.xlsx("Investment Analysis Tool.xlsx", sheetName = "funds")
tickers <- funds %>% pull(ticker)



# # Historical Prices
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



# year to date & 6 month rolling

six_months <- today() - months(6)
year_start <- as.Date("2020-01-01")

earliest_date <- min(six_months, year_start)
today <- today()

stocks <- tq_get(tickers,
                 from = earliest_date,
                 to = today,
                 get = "stock.prices")

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

returns <- bind_rows(ytd_return, sixmo_return) %>%
  pivot_wider(names_from = type, values_from = return) %>%
  ungroup() %>%
  mutate(rank_ytd = rank(desc(ytd)),
         rank_6mo = rank(desc(rolling_6_mo))) %>%
  arrange(rank_ytd)

percent_year <- as.numeric(today - year_start)/365

returns_full <- full_join(funds, returns, by = "ticker") %>% 
  mutate(ytd_adj = ytd - expense_ratio * percent_year) %>%
  mutate(rank_ytd = rank(desc(ytd_adj))) %>%
  arrange(rank_ytd)

returns_class <- returns_full %>%
  group_by(asset_class, size, type, location) %>%
  summarize(avg_ytd = mean(ytd_adj),
            avg_6mo = mean(rolling_6_mo),
            funds = n()) %>%
  arrange(desc(avg_ytd))



returns <- data.frame(returns)
removeSheet(wb, sheetName = "ytd returns")
newSheet <- createSheet(wb, sheetName="ytd returns")
addDataFrame(returns, newSheet, row.names = F)

saveWorkbook(wb, "Investment Analysis Tool.xlsx")