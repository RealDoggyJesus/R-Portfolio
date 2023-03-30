#Packages needed for this...
library("tidyverse")
library(openxlsx)
library("readxl")
set.seed(123)

#Loading dataset
df <- read.csv("PlaidPantryBig.csv")
df$Date <- as.Date(df$Date)

#Omit irrelevant rows to make the loading and calc easier
df <- df[,c(-2:-8, -17:-29)]

#Switch Dec 16 on sales values to ly sales values
df_ly <- df

df_ly$Units.LY <- ifelse(df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-12-16")) & df_ly$Date <= 
                               as.Date(paste0(year(df_ly$Date), "-12-31")), df_ly$Units, df_ly$Units.LY)
df_ly$X..Sales.LY <- ifelse(df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-12-16")) & df_ly$Date <= 
                              as.Date(paste0(year(df_ly$Date), "-12-31")), df_ly$X..Sales, df_ly$X..Sales.LY)

df_ly <- df_ly[,c(-6,-8)]

df_ly$Date <- df_ly$Date %m-% years(1)

colnames(df_ly)[6] = "Sales"
colnames(df_ly)[7] = "Qty Sold"

# Format the dates to get the week ending date (week ending on Thursday)
#This is only good for 2021
df_ly$week_ending <- df_ly$Date + 4 - as.numeric(format(df_ly$Date, "%w"))
df_ly$week_ending[df_ly$week_ending < df_ly$Date] <- df_ly$week_ending[df_ly$week_ending < df_ly$Date] + 7

# Adjust the week ending date for the first week of the year
df_ly$week_ending[df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-01-01")) & 
                    df_ly$Date <= as.Date(paste0(year(df_ly$Date), "-01-07"))] <- 
  as.Date(paste0(year(df_ly$Date), "-01-07"))

# Adjust the week ending date for the last week of the year
df_ly$week_ending[df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-12-24")) & 
                    df_ly$Date <= as.Date(paste0(year(df_ly$Date), "-12-30"))] <- 
  as.Date(paste0(year(df_ly$Date) , "-12-30"))

# Adjust the week ending date for the last week of the previous year
df_ly$week_ending[df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-12-31"))] <- 
  as.Date(paste0(year(df_ly$Date) + 1, "-01-08")) - 2

###------------------------------------------------------------###
#Creating a week number column based off of the date in the date column
# Format the dates to get the week number (week starting on Friday)
#This is only good for 2021
df_ly$week <- as.numeric(format(df_ly$Date + 2, "%U"))

# Adjust the week number for the first week of the year
df_ly$week[df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-01-01")) & 
             +           df_ly$Date <= as.Date(paste0(year(df_ly$Date), "-01-05"))] <- 1

# Adjust the week number for the last week of the year
df_ly$week[df_ly$Date >= as.Date(paste0(year(df_ly$Date), "-12-31"))] <- 1

# Adjust the week number for the last week of the previous year
df_ly$week[df_ly$Date <= as.Date(paste0(year(df_ly$Date) - 1, "-12-30")) & 
             +           df_ly$Date >= as.Date(paste0(year(df_ly$Date) - 1, "-12-25"))] <- 52

#change to last year

df_lys <- df_ly[,-7]
df_lyu <- df_ly[,-6]

colnames(df_lys)[1] = "OrgGroupingValueItemUPCCY2"
colnames(df_lys)[2] = "ProdGroupingValueItemUPCCY"
colnames(df_lys)[3] = "ItemDescriptionUPCCY2"
colnames(df_lys)[4] = "UPCValueCY2"
colnames(df_lys)[5] = "DayDateorgAndProdUPCCY"
colnames(df_lys)[6] = "MeasureValueOrgAndProdUPCCY"
colnames(df_lys)[7] = "WeekEnding"
colnames(df_lys)[8] = "WeekNumber"


df_lys["MeasureLabelOrgAndProdUPCCY"] <- "Sales"

df_lys <- df_lys[c("OrgGroupingValueItemUPCCY2", "ProdGroupingValueItemUPCCY", "ItemDescriptionUPCCY2", "UPCValueCY2","DayDateorgAndProdUPCCY", "MeasureLabelOrgAndProdUPCCY", "MeasureValueOrgAndProdUPCCY", "WeekEnding", "WeekNumber")]

colnames(df_lyu)[1] = "OrgGroupingValueItemUPCCY2"
colnames(df_lyu)[2] = "ProdGroupingValueItemUPCCY"
colnames(df_lyu)[3] = "ItemDescriptionUPCCY2"
colnames(df_lyu)[4] = "UPCValueCY2"
colnames(df_lyu)[5] = "DayDateorgAndProdUPCCY"
colnames(df_lyu)[6] = "MeasureValueOrgAndProdUPCCY"
colnames(df_lyu)[7] = "WeekEnding"
colnames(df_lyu)[8] = "WeekNumber"

df_lyu["MeasureLabelOrgAndProdUPCCY"] <- "Qty Sold"

df_lyu <- df_lyu[c("OrgGroupingValueItemUPCCY2", "ProdGroupingValueItemUPCCY", "ItemDescriptionUPCCY2", "UPCValueCY2","DayDateorgAndProdUPCCY", "MeasureLabelOrgAndProdUPCCY", "MeasureValueOrgAndProdUPCCY", "WeekEnding", "WeekNumber")]

df_lys <- subset(df_lys, df_lys$MeasureValueOrgAndProdUPCCY != 0)
df_lyu <- subset(df_lyu, df_lyu$MeasureValueOrgAndProdUPCCY != 0)

#working with the current year dataframe
minidf <- subset(df, df$Date <= as.Date(paste0(year(df$Date), "-12-15")))
minidf <- minidf[,c(-7,-9)]
colnames(minidf)[6] = "Sales"
colnames(minidf)[7] = "Qty Sold"

week51s <- read_excel("Week 51 2022.xlsx", sheet = "wk 51 $")
week51s <- week51s[,c(-8:-10)]
week51u <- read_excel("Week 51 2022.xlsx", sheet = "wk 51 u")
week51u <- week51u[,c(-8:-10)]

week52s <- read_excel("Week 52.xlsx", sheet = "wk 52 $")
week52s["MeasureLabelOrgAndProdUPCCY"] <- "Sales"
colnames(week52s)[6] = "MeasureValueOrgAndProdUPCCY"
week52s <- week52s[c("OrgGroupingValueItemUPCCY2", "ProdGroupingValueItemUPCCY", "ItemDescriptionUPCCY2", "UPCValueCY2","DayDateorgAndProdUPCCY", "MeasureLabelOrgAndProdUPCCY", "MeasureValueOrgAndProdUPCCY")]

week52u <- read_excel("Week 52.xlsx", sheet = "wk 52 u")
week52u["MeasureLabelOrgAndProdUPCCY"] <- "Qty Sold"
colnames(week52u)[6] = "MeasureValueOrgAndProdUPCCY"
week52u <- week52u[c("OrgGroupingValueItemUPCCY2", "ProdGroupingValueItemUPCCY", "ItemDescriptionUPCCY2", "UPCValueCY2","DayDateorgAndProdUPCCY", "MeasureLabelOrgAndProdUPCCY", "MeasureValueOrgAndProdUPCCY")]

minidfs <- minidf[,-7]
minidfu <- minidf[,-6]

colnames(minidfs)[1] = "OrgGroupingValueItemUPCCY2"
colnames(minidfs)[2] = "ProdGroupingValueItemUPCCY"
colnames(minidfs)[3] = "ItemDescriptionUPCCY2"
colnames(minidfs)[4] = "UPCValueCY2"
colnames(minidfs)[5] = "DayDateorgAndProdUPCCY"
colnames(minidfs)[6] = "MeasureValueOrgAndProdUPCCY"

minidfs["MeasureLabelOrgAndProdUPCCY"] <- "Sales"

minidfs <- minidfs[c("OrgGroupingValueItemUPCCY2", "ProdGroupingValueItemUPCCY", "ItemDescriptionUPCCY2", "UPCValueCY2","DayDateorgAndProdUPCCY", "MeasureLabelOrgAndProdUPCCY", "MeasureValueOrgAndProdUPCCY")]

colnames(minidfu)[1] = "OrgGroupingValueItemUPCCY2"
colnames(minidfu)[2] = "ProdGroupingValueItemUPCCY"
colnames(minidfu)[3] = "ItemDescriptionUPCCY2"
colnames(minidfu)[4] = "UPCValueCY2"
colnames(minidfu)[5] = "DayDateorgAndProdUPCCY"
colnames(minidfu)[6] = "MeasureValueOrgAndProdUPCCY"

minidfu["MeasureLabelOrgAndProdUPCCY"] <- "Qty Sold"

minidfu <- minidfu[c("OrgGroupingValueItemUPCCY2", "ProdGroupingValueItemUPCCY", "ItemDescriptionUPCCY2", "UPCValueCY2","DayDateorgAndProdUPCCY", "MeasureLabelOrgAndProdUPCCY", "MeasureValueOrgAndProdUPCCY")]

minidfs <- rbind(week51s, week52s, minidfs)
minidfs <- subset(minidfs, minidfs$MeasureValueOrgAndProdUPCCY != 0)

#Creating Minidf week ending and week number
##minidf$week_ending <- floor_date(as.numeric(minidf$date, unit = "week", week_start = 5) + days(6))
minidfs$WeekEnding <- floor_date(minidfs$DayDateorgAndProdUPCCY, unit = "week", week_start = 5) + days(6)

###------------------------------------------------------------###
# Format the dates to get the week number (week starting on Friday)
minidfs$WeekNumber <- as.numeric(format(minidfs$DayDateorgAndProdUPCCY + 2, "%U"))

# Adjust the week number for the first week of the year
minidfs$WeekNumber[minidfs$DayDateorgAndProdUPCCY >= as.Date(paste0(year(minidfs$DayDateorgAndProdUPCCY), "-01-01")) & 
          +           minidfs$DayDateorgAndProdUPCCY <= as.Date(paste0(year(minidfs$DayDateorgAndProdUPCCY), "-01-06"))] <- 1

minidfu <- rbind(week51u, week52u, minidfu)
minidfu <- subset(minidfu, minidfu$MeasureValueOrgAndProdUPCCY != 0)

#Creating Minidf week ending and week number
##minidf$week_ending <- floor_date(as.numeric(minidf$date, unit = "week", week_start = 5) + days(6))
minidfu$WeekEnding <- floor_date(minidfu$DayDateorgAndProdUPCCY, unit = "week", week_start = 5) + days(6)

###_____________________________###


minidfu$WeekNumber <- mapply(function(date) {
  year_start <- as.Date(paste0(format(date, "%Y"), "-01-01"))
  week_num <- isoweek(floor_date(date, unit = "week", week_start = 5))
  week_diff <- week_num - isoweek(year_start) + 1
  if (week_diff < 1) {
    week_diff <- week_diff + 52
  }
  week_diff
}, minidfu$DayDateorgAndProdUPCCY)


#importing data to append

week1s <- read_excel("Week 1 2023.xlsx", sheet = "wk 1 $")
week1u <- read_excel("Week 1 2023.xlsx", sheet = "wk 1 u")

week2s <- read_excel("Week 2 2023.xlsx", sheet = "wk 2 $")
week2u <- read_excel("Week 2 2023.xlsx", sheet = "wk 2 u")

week3s <- read_excel("Week 3 2023.xlsx", sheet = "wk 3 $")
week3u <- read_excel("Week 3 2023.xlsx", sheet = "wk 3 u")
  
week4s <- read_excel("Week 4 2023.xlsx", sheet = "wk 4 $")
week4u <- read_excel("Week 4 2023.xlsx", sheet = "wk 4 u")

week5s <- read_excel("Week 5 2023.xlsx", sheet = "wk 5 $")
week5u <- read_excel("Week 5 2023.xlsx", sheet = "wk 5 u")

week6s <- read_excel("Week 6 2023.xlsx", sheet = "wk 6 $")
week6u <- read_excel("Week 6 2023.xlsx", sheet = "wk 6 u")

week7s <- read_excel("Week 7 2023.xlsx", sheet = "wk 7 $")
week7u <- read_excel("Week 7 2023.xlsx", sheet = "wk 7 u")

week8s <- read_excel("Week 8 2023.xlsx", sheet = "wk 8 $")
week8u <- read_excel("Week 8 2023.xlsx", sheet = "wk 8 u")

week9s <- read_excel("Week 9 2023.xlsx", sheet = "wk 9 $")
week9u <- read_excel("Week 9 2023.xlsx", sheet = "wk 9 u")

#combining units and sales tables
df23s <- rbind(week1s, week2s, week3s, week4s, week5s, week6s, week7s, week8s, week9s)
df23s <- subset(df23s, df23s["MeasureValueOrgAndProdUPCCY"] != 0)

df23u <- rbind(week1u, week2u, week3u, week4u, week5u, week6u, week7u, week8u, week9u)
df23u <- subset(df23u, df23u["MeasureValueOrgAndProdUPCCY"] != 0)

#Need to complete the week ending and week number logic here.

df23u$WeekEnding <- floor_date(df23u$DayDateorgAndProdUPCCY, unit = "week", week_start = 5) + days(6)

df23u$WeekNumber <- mapply(function(date) {
  year_start <- as.Date(paste0(format(date, "%Y"), "-01-01"))
  week_num <- isoweek(floor_date(date, unit = "week", week_start = 5))
  week_diff <- week_num - isoweek(year_start) + 1
  if (week_diff < 1) {
    week_diff <- week_diff + 52
  }
  week_diff
}, df23u$DayDateorgAndProdUPCCY)

df23s$WeekEnding <- floor_date(df23s$DayDateorgAndProdUPCCY, unit = "week", week_start = 5) + days(6)

df23s$WeekNumber <- mapply(function(date) {
  year_start <- as.Date(paste0(format(date, "%Y"), "-01-01"))
  week_num <- isoweek(floor_date(date, unit = "week", week_start = 5))
  week_diff <- week_num - isoweek(year_start) + 1
  if (week_diff < 1) {
    week_diff <- week_diff + 52
  }
  week_diff
}, df23s$DayDateorgAndProdUPCCY)

#Combine all of the sales data here into one master file
totalsales <- rbind(df23s, minidfs, df_lys)
  
#Combine all of the units data here into one master file
totalunits <- rbind(df23u, minidfu, df_lyu)
  
#Group the tables to eliminate duplicates
totalsales <- totalsales %>%
  group_by(OrgGroupingValueItemUPCCY2, ProdGroupingValueItemUPCCY, ItemDescriptionUPCCY2, UPCValueCY2, DayDateorgAndProdUPCCY, MeasureLabelOrgAndProdUPCCY, WeekNumber, WeekEnding) %>%
  summarise(across(MeasureValueOrgAndProdUPCCY, sum))

totalunits <- totalunits %>%
  group_by(OrgGroupingValueItemUPCCY2, ProdGroupingValueItemUPCCY, ItemDescriptionUPCCY2, UPCValueCY2, DayDateorgAndProdUPCCY, MeasureLabelOrgAndProdUPCCY, WeekNumber, WeekEnding) %>%
  summarise(across(MeasureValueOrgAndProdUPCCY, sum))

rm(df)
rm(df_ly)
rm(df_lyu)
rm(df_lys)
rm(df23s)
rm(df23u)
rm(minidfs)
rm(minidfu)
rm(minidf)

write.csv(totalsales, "C:\\Python39\\PlaidPantryAggS.csv", row.names=FALSE)

write.csv(totalunits, "C:\\Python39\\PlaidPantryAggU.csv", row.names=FALSE)

#Creating the excel doc
dataset_names <- list('21-23 U' = totalunits, '21-23 S' = totalsales)
write.xlsx(dataset_names, file = 'PlaidPantry 2021-2023.xlsx')


#working with the current year dataframe
#df <- subset(df, df$Date <= as.Date(paste0(year(df$Date), "-12-15")))
#df <- df[,c(-7,-9)]
#colnames(df)[6] = "Sales"
#colnames(df)[7] = "Qty Sold"

#pivot <- total %>%
#  pivot_longer(cols = c("Dollars", "Qty Sold"),
#               names_to = "MeasureLabelOrgAndProdUPCCY",
#               values_to = "MeasureValueOrgAndProdUPCCY")

#colnames(pivot)[1] = "OrgGroupingValueItemUPCCY2"
#colnames(pivot)[2] = "ProdGroupingValueItemUPCCY"
#colnames(pivot)[3] = "ItemDescriptionUPCCY2"
#colnames(pivot)[4] = "UPCValueCY2"
#colnames(pivot)[5] = "DayDateorgAndProdUPCCY"

#pivot <- pivot %>%
#  group_by(OrgGroupingValueItemUPCCY2, ProdGroupingValueItemUPCCY, ItemDescriptionUPCCY2, UPCValueCY2, DayDateorgAndProdUPCCY, MeasureLabelOrgAndProdUPCCY) %>%
#  summarise(across(MeasureValueOrgAndProdUPCCY, sum))

#Export the CSV
#write.csv(pivot, "C:\\Python39\\PlaidPantryAgg.csv", row.names=FALSE)