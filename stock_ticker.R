# File:   stock_ticker.R - R program
# Author: Rish Singhania
# Additional Editors: Zach Bethune, Kellie Forrester, Brian Thomas,
# Travis Cyronek, Ryan Sherrard, Sarah Papich
# Last edit: 4.01.22

# Purpose: Calculates 1 week, 4 week, and 52 week changes for a list of stocks
# specified in the master database, and outputs the results in CSV files. The
# program also computes certain stock indexes and writes the results in an
# Excel file.


#----------------------------------------4---------------------------
# Housekeeping
#-------------------------------------------------------------------

rm(list=ls())
# Load required libraries
library("quantmod")
library("xts")
library("reshape")
library("writexl")
library("scales")
library("plyr")
library("conflicted")
library("tidyverse")
library("magrittr")
conflicts_prefer(xts::last)
conflicts_prefer(dplyr::filter)
library(gmailr) #to send the email
working.dir <- "/Users/jonahdanziger/Library/CloudStorage/Box-Box/efp customer facing/requests (free)/pcbt (stock ticker)"
setwd(working.dir)

# -------------------------------------------------------------------------


# -------------------------------------------------------------------------



#-------------------------------------------------------------------
# Read list of stocks from the master stock file.
#-------------------------------------------------------------------

stocklist <- read.csv(paste(getwd(),"/stock_list.csv",sep=""), head=TRUE, sep=",", colClasses=c("character", "character"))
#sum

#remove ATVI, which produces an error
stocklist %<>%
  filter(ticker!="ATVI")
stocklist %<>%
  filter(ticker!="CAMP")
stocklist %<>%
  filter(ticker!="GPS")
#Juniper Networks (JNPR) was sold on July 2, 2025 so no longer has a ticker
stocklist %<>%
  filter(ticker!="JNPR")
#PPBI stock no longer trades under that symbol. On August 31, 2025, the acquisition of Pacific Premier Bancorp, Inc. by Columbia Banking System, Inc. was completed
stocklist %<>%
  filter(ticker!="PPBI")
stocklist %<>%
  filter(ticker!="CVGW")
#-------------------------------------------------------------------
# Read stock data using getSymbols. Need to delete the ^ symbol after
# downloading data from the index tickers; R doesn't read that symbol in
# variable names.
#-------------------------------------------------------------------

#getSymbols(stocklist$ticker[1], verbose=TRUE, src="FRED")

start<-1
lastcompletedloop<-1
#skipped:   (once: )

#setInternet2(use = TRUE)
if (lastcompletedloop < length(stocklist$ticker)) {
    print('Not Done')
    start <- lastcompletedloop
}
for(i in  start:length(stocklist$ticker)){
    getSymbols(stocklist$ticker[i], verbose=TRUE, src="yahoo")
    lastcompletedloop <- i
}

stocklist$ticker <- gsub("\\^", "", stocklist$ticker)


#-------------------------------------------------------------------
# Determine most recent end of week date, drop stocks that stop reporting.
#-------------------------------------------------------------------

wklydata <- to.weekly(get("IXIC"))
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
NASDAQdata <- wklydata
mostrecent_endofweek <- last(index(NASDAQdata))


#-------------------------------------------------------------------
# Construct pcbt, ca, tricounty indexes
# To stay consistent with how I read the rest of the stocks I will create tickers for these indices.
# The pcbt index with have the ticker PCBTEFP
# The CA index will have the ticker CAEFP
# The tricounty index will have the ticker TCEFP
# I will add these tickers to the existing stocklist$tickers and put 1s at the us.index.
#-------------------------------------------------------------------

PCBTEFP <- 1
TCEFP <- 0
CAEFP <- 0

PCBTEFP_1week <- 0
TCEFP_1week <- 0
CAEFP_1week <- 0

PCBTEFP_4week <- 0
TCEFP_4week <- 0
CAEFP_4week <- 0

PCBTEFP_52week <- 0
TCEFP_52week <- 0
CAEFP_52week <- 0

for(i in 1:length(stocklist$ticker)){
    if (stocklist$pcbt.index[i] == 1){
        PCBTEFP <- PCBTEFP + get(stocklist$ticker[i])
    }
    if (stocklist$tri.county.index[i] == 1){
        TCEFP <- TCEFP + get(stocklist$ticker[i])
    }
    if (stocklist$ca.index[i] == 1){
        CAEFP <- CAEFP + get(stocklist$ticker[i])
    }
    if (index(get(stocklist$ticker[i]))[1]<seq(mostrecent_endofweek, length=2,by="-1 weeks")[2]){
        if (stocklist$pcbt.index[i] == 1){
            PCBTEFP_1week <- PCBTEFP_1week + get(stocklist$ticker[i])
        }
        if (stocklist$tri.county.index[i] == 1){
            TCEFP_1week <- TCEFP_1week + get(stocklist$ticker[i])
        }
        if (stocklist$ca.index[i] == 1){
            CAEFP_1week <- CAEFP_1week + get(stocklist$ticker[i])
        }
    }
    if (index(get(stocklist$ticker[i]))[1]<seq(mostrecent_endofweek, length=2,by="-4 weeks")[2]){
        if (stocklist$pcbt.index[i] == 1){
            PCBTEFP_4week <- PCBTEFP_4week + get(stocklist$ticker[i])
        }
        if (stocklist$tri.county.index[i] == 1){
            TCEFP_4week <- TCEFP_4week + get(stocklist$ticker[i])
        }
        if (stocklist$ca.index[i] == 1){
            CAEFP_4week <- CAEFP_4week + get(stocklist$ticker[i])
        }
    }
    if (index(get(stocklist$ticker[i]))[1]<seq(mostrecent_endofweek, length=2,by="-52 weeks")[2]){
        if (stocklist$pcbt.index[i] == 1){
            PCBTEFP_52week <- PCBTEFP_52week + get(stocklist$ticker[i])
        }
        if (stocklist$tri.county.index[i] == 1){
            TCEFP_52week <- TCEFP_52week + get(stocklist$ticker[i])
        }
        if (stocklist$ca.index[i] == 1){
            CAEFP_52week <- CAEFP_52week + get(stocklist$ticker[i])
        }
    }
}

pcbt.row <- c("PCBTEFP", "Pacific Coast Business Times Stock Index", "0", "0", "0", "1")
tri.county.row <- c("TCEFP", "Tri County Stock Index", "0", "0", "0", "1")
ca.row <- c("CAEFP", "California Stock Index", "0", "0", "0", "1")

stocklist <- rbind(stocklist, pcbt.row, tri.county.row, ca.row)


#-------------------------------------------------------------------
# Convert stock data to weekly data. This conversion is accomplished using the
# to.weekly function. This function converts daily stock price data into weekly
# data simply by changing the observation time period to one week. That is, it
# reports the opening price of a stock on Monday and the closing price on
# Friday, or the last available day of the week.

# Since we are interested in calculating Friday-to-Friday changes. The code
# below drops all data after the last Friday in the data set. I accomplish this
# by first reading the data into a local variable "wklydata", finding the name
# of the weekdays that this data is indexed using (note the xts library indexes
# time series using dates in this case, so index(wklydata) returns the date.)
# Historically, the stock market has closed either on Thursdays or on Fridays
# so I check whether the last element in the dataset equals either of these
# days.

# CAUTION: This means that the program must be run before Wednesday to get last
# week's information.

# The assign function allows us to create new variables using strings. See R
# help.
#-------------------------------------------------------------------

for(i in 1:length(stocklist$ticker)) {
    wklydata <- to.weekly(get(stocklist$ticker[i]))
    wkdaylist <- weekdays(index(wklydata))
    if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
        wklydata <- wklydata[-length(Cl(wklydata)),]
    }
    assign(paste(stocklist$ticker[i], "wklydata", sep="."), wklydata)
}


#-------------------------------------------------------------------
# Calculate 52 week change for indicies, kicking out stocks that have had IPO
# dates within the last year.
#-------------------------------------------------------------------

indicies <- c("PCBTEFP", "TCEFP", "CAEFP")
index_1week_pctchg <- rep(0,length(indicies))
index_4week_pctchg <- rep(0,length(indicies))
index_52week_pctchg <- rep(0,length(indicies))

#PCBTEFP Index
wklydata<- to.weekly(PCBTEFP_1week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_1week_pctchg[1] <- last(Delt(wklydata$PCBTEFP_1week.Close, k=1, type="arithmetic") * 100)

wklydata<- to.weekly(PCBTEFP_4week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_4week_pctchg[1] <- last(Delt(wklydata$PCBTEFP_4week.Close, k=4, type="arithmetic") * 100)

wklydata<- to.weekly(PCBTEFP_52week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_52week_pctchg[1] <- last(Delt(wklydata$PCBTEFP_52week.Close, k=52, type="arithmetic") * 100)

#TCEFP Index
wklydata<- to.weekly(TCEFP_1week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_1week_pctchg[2] <- last(Delt(wklydata$TCEFP_1week.Close, k=1, type="arithmetic") * 100)

wklydata<- to.weekly(TCEFP_4week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_4week_pctchg[2] <- last(Delt(wklydata$TCEFP_4week.Close, k=4, type="arithmetic") * 100)

wklydata<- to.weekly(TCEFP_52week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_52week_pctchg[2] <- last(Delt(wklydata$TCEFP_52week.Close, k=52, type="arithmetic") * 100)

#CAEFP Index

wklydata<- to.weekly(CAEFP_1week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_1week_pctchg[3] <- last(Delt(wklydata$CAEFP_1week.Close, k=1, type="arithmetic") * 100)

wklydata<- to.weekly(CAEFP_4week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_4week_pctchg[3] <- last(Delt(wklydata$CAEFP_4week.Close, k=4, type="arithmetic") * 100)

wklydata<- to.weekly(CAEFP_52week)
wkdaylist <- weekdays(index(wklydata))
if (last(wkdaylist) != "Friday" && last(wkdaylist) != "Thursday") {
    wklydata <- wklydata[-length(Cl(wklydata)),]
}
index_52week_pctchg[3] <- last(Delt(wklydata$CAEFP_52week.Close, k=52, type="arithmetic") * 100)

#-------------------------------------------------------------------
# Calculate n1.week.change, n4.week.change, and n52.week.change for all stocks.
# I save this data in the stocklist frame using these variable names.

# The process for calculating the n week change is as follows. First I check to
# see if we can calculate the n week change. For example, this calculation
# might not be possible for companies that have been public for less than n
# weeks.

# Second I use the "Delt" function on the closing prices (found using the "Cl"
# function) to calculate the n week changes, and then I pick the last value of
# these changes (using the function "last")

# If n week changes cannot be calculated. I assign the value NA.

# At the end I get the latest quote of the stock
#-------------------------------------------------------------------


# Create the columns in the stocklist data frame
stocklist$n1.week.change <- NA
stocklist$n4.week.change <- NA
stocklist$n52.week.change <- NA
stocklist$latest.quote <- NA

# Loop through to calculate 1 week, 4 week, and 52 week percentage changes
for(i in 1:length(stocklist$ticker)) {
    # Define useful variables
    stock.close.prices <- Cl(get(paste(stocklist$ticker[i], "wklydata", sep=".")))
    n1.var <- paste(stocklist$ticker[i], "n1.week.change", sep=".")
    n4.var <- paste(stocklist$ticker[i], "n4.week.change", sep=".")
    n52.var <- paste(stocklist$ticker[i], "n52.week.change", sep=".")

    #-------------------------------------------------------------------
    # 1 week change
    if(length(stock.close.prices) > 1) {
        assign(n1.var, last(Delt(stock.close.prices, type="arithmetic") * 100))
    }
    if(length(stock.close.prices) <= 1 & stocklist$ticker[i] %in% indicies) {
        assign(n1.var, index_1week_pctchg[which(indicies == stocklist$ticker[i])])
    }
    if(length(stock.close.prices) <= 1 & !(stocklist$ticker[i] %in% indicies)) {
        assign(n1.var, NA)
    }
    # Save n1.week.change in stocklist data frame
    stocklist$n1.week.change[i] <- get(n1.var)

    #-------------------------------------------------------------------
    # 4 week change
    if(length(stock.close.prices) > 4) {
        assign(n4.var, last(Delt(stock.close.prices, k=4, type="arithmetic") * 100))
    }
    if(length(stock.close.prices) <= 4 & stocklist$ticker[i] %in% indicies) {
        assign(n4.var, index_4week_pctchg[which(indicies == stocklist$ticker[i])])
    }
    if(length(stock.close.prices) <= 4 & !(stocklist$ticker[i] %in% indicies)) {
        assign(n4.var, NA)
    }
    # Save n4.week.change in a data frame
    stocklist$n4.week.change[i] <- get(n4.var)

    #-------------------------------------------------------------------
    # 52 week change
    if(length(stock.close.prices) > 52) {
        assign(n52.var, last(Delt(stock.close.prices, k=52, type="arithmetic") * 100))
    }
    if(length(stock.close.prices) <= 52 & stocklist$ticker[i] %in% indicies) {
        assign(n52.var, index_52week_pctchg[which(indicies == stocklist$ticker[i])])
    }
    if(length(stock.close.prices) <= 52 & !(stocklist$ticker[i] %in% indicies)) {
        assign(n52.var, NA)
    }
    # Save 52wkchng in a data frame
    stocklist$n52.week.change[i] <- get(n52.var)
    #-------------------------------------------------------------------

    # Save latest quote
    stocklist$latest.quote[i] <- last(stock.close.prices)
}


#-------------------------------------------------------------------
# Calculate top 5 winners and losers. Need to omit stock indices for this
# calculation. I drop the NAs when sorting.
#-------------------------------------------------------------------

stocklist.nat <- stocklist[stocklist$us.index==1, ]
stocklist.pcbt <- stocklist[stocklist$pcbt.index==1, ]
stocklist.ca <- stocklist[stocklist$ca.index==1, ]
stocklist.tricounty <- stocklist[stocklist$tri.county.index==1, ]

# Find winners and losers
#stocklist.dropindexes <- stocklist[stocklist$us.index==0, ]
stocklist.sortedwinners <- stocklist.pcbt[order(stocklist.pcbt$n1.week.change, decreasing=TRUE, na.last=NA), ]
stocklist.sortedlosers <- stocklist.pcbt[order(stocklist.pcbt$n1.week.change, decreasing=FALSE, na.last=NA), ]
stocklist.top5winners <- head(stocklist.sortedwinners, n=5)
stocklist.top5losers <- head(stocklist.sortedlosers, n=5)


stocklist.top5losers<-stocklist.top5losers[, c("ticker", "company.name", "n1.week.change", "n4.week.change", "n52.week.change", "latest.quote")]
stocklist.top5winners<-stocklist.top5winners[, c("ticker", "company.name", "n1.week.change", "n4.week.change", "n52.week.change", "latest.quote")]
stocklist.pcbt<-stocklist.pcbt[, c("ticker", "company.name", "n1.week.change", "n4.week.change", "n52.week.change", "latest.quote")]
stocklist.nat<-stocklist.nat[, c("ticker", "company.name", "n1.week.change", "n4.week.change", "n52.week.change", "latest.quote")]


#-------------------------------------------------------------------
# Get date of last Friday. This will be the final file name. I use the S&P
# index to get the last date. Filename format is "PCBT_year_mm_dd"
#-------------------------------------------------------------------

fname <- paste("PCBT", gsub("-", "_", last(index(GSPC.wklydata))), sep="_")
fname <- paste(fname, "xlsx", sep=".")
setwd(paste(getwd(),"/Weekly_Reports",sep=""))


####### Make into data frame and rename the columns

stocklist.top5winners<-data.frame(stocklist.top5winners)
stocklist.top5winners<-reshape::rename(stocklist.top5winners,c(ticker="Symbol", company.name="Company Name",n1.week.change="Change from Previous Week", n4.week.change="Change from Last Four Weeks", n52.week.change="Change from Last 52 Weeks",latest.quote="Latest Quote"))
stocklist.top5losers<-data.frame(stocklist.top5losers)
stocklist.top5losers<-reshape::rename(stocklist.top5losers,c(ticker="Symbol", company.name="Company Name",n1.week.change="Change from Previous Week", n4.week.change="Change from Last Four Weeks", n52.week.change="Change from Last 52 Weeks",latest.quote="Latest Quote"))

stocklist.pcbt<-data.frame(stocklist.pcbt)
stocklist.pcbt<-reshape::rename(stocklist.pcbt,c(ticker="Symbol", company.name="Company Name",n1.week.change="Change from Previous Week", n4.week.change="Change from Last Four Weeks", n52.week.change="Change from Last 52 Weeks",latest.quote="Latest Quote"))

stocklist.nat<-data.frame(stocklist.nat)
stocklist.nat<-reshape::rename(stocklist.nat,c(ticker="Symbol", company.name="Company Name",n1.week.change="Change from Previous Week", n4.week.change="Change from Last Four Weeks", n52.week.change="Change from Last 52 Weeks",latest.quote="Latest Quote"))


####### Change to Numeric and Round to two decimal places

stocklist.top5losers$'Change from Previous Week'<-as.numeric(stocklist.top5losers$'Change from Previous Week')
stocklist.top5losers$'Change from Previous Week'<-format(round(stocklist.top5losers$'Change from Previous Week', 2), nsmall = 2)
stocklist.top5losers$'Change from Previous Week'<-as.numeric(stocklist.top5losers$'Change from Previous Week')
stocklist.top5losers$'Change from Previous Week'<-sprintf("%1.2f%%", stocklist.top5losers$'Change from Previous Week')
stocklist.top5losers$'Change from Last Four Weeks'<-as.numeric(stocklist.top5losers$'Change from Last Four Weeks')
stocklist.top5losers$'Change from Last Four Weeks'<-format(round(stocklist.top5losers$'Change from Last Four Weeks', 2), nsmall = 2)
stocklist.top5losers$'Change from Last Four Weeks'<-as.numeric(stocklist.top5losers$'Change from Last Four Weeks')
stocklist.top5losers$'Change from Last Four Weeks'<-sprintf("%1.2f%%", stocklist.top5losers$'Change from Last Four Weeks')
stocklist.top5losers$'Change from Last 52 Weeks'<-as.numeric(stocklist.top5losers$'Change from Last 52 Weeks')
stocklist.top5losers$'Change from Last 52 Weeks'<-format(round(stocklist.top5losers$'Change from Last 52 Weeks', 2), nsmall = 2)
stocklist.top5losers$'Change from Last 52 Weeks'<-as.numeric(stocklist.top5losers$'Change from Last 52 Weeks')
stocklist.top5losers$'Change from Last 52 Weeks'<-sprintf("%1.2f%%", stocklist.top5losers$'Change from Last 52 Weeks')

stocklist.top5winners$'Change from Previous Week'<-as.numeric(stocklist.top5winners$'Change from Previous Week')
stocklist.top5winners$'Change from Previous Week'<-format(round(stocklist.top5winners$'Change from Previous Week', 2), nsmall = 2)
stocklist.top5winners$'Change from Previous Week'<-as.numeric(stocklist.top5winners$'Change from Previous Week')
stocklist.top5winners$'Change from Previous Week'<-sprintf("%1.2f%%", stocklist.top5winners$'Change from Previous Week')
stocklist.top5winners$'Change from Last Four Weeks'<-as.numeric(stocklist.top5winners$'Change from Last Four Weeks')
stocklist.top5winners$'Change from Last Four Weeks'<-format(round(stocklist.top5winners$'Change from Last Four Weeks', 2), nsmall = 2)
stocklist.top5winners$'Change from Last Four Weeks'<-as.numeric(stocklist.top5winners$'Change from Last Four Weeks')
stocklist.top5winners$'Change from Last Four Weeks'<-sprintf("%1.2f%%", stocklist.top5winners$'Change from Last Four Weeks')
stocklist.top5winners$'Change from Last 52 Weeks'<-as.numeric(stocklist.top5winners$'Change from Last 52 Weeks')
stocklist.top5winners$'Change from Last 52 Weeks'<-format(round(stocklist.top5winners$'Change from Last 52 Weeks', 2), nsmall = 2)
stocklist.top5winners$'Change from Last 52 Weeks'<-as.numeric(stocklist.top5winners$'Change from Last 52 Weeks')
stocklist.top5winners$'Change from Last 52 Weeks'<-sprintf("%1.2f%%", stocklist.top5winners$'Change from Last 52 Weeks')

stocklist.top5winners$'Latest Quote'<-as.numeric(stocklist.top5winners$'Latest Quote')
stocklist.top5winners$'Latest Quote'<-dollar(stocklist.top5winners$'Latest Quote')

stocklist.top5losers$'Latest Quote'<-as.numeric(stocklist.top5losers$'Latest Quote')
stocklist.top5losers$'Latest Quote'<-dollar(stocklist.top5losers$'Latest Quote')

############   Do same for PCBT sheet
############   Change to Numeric and Round to two decimal places

stocklist.pcbt$'Change from Previous Week'<-as.numeric(stocklist.pcbt$'Change from Previous Week')
stocklist.pcbt$'Change from Previous Week'<-format(round(stocklist.pcbt$'Change from Previous Week', 2), nsmall = 2)
stocklist.pcbt$'Change from Previous Week'<-as.numeric(stocklist.pcbt$'Change from Previous Week')
stocklist.pcbt$'Change from Previous Week'<-sprintf("%1.2f%%", stocklist.pcbt$'Change from Previous Week')
stocklist.pcbt$'Change from Last Four Weeks'<-as.numeric(stocklist.pcbt$'Change from Last Four Weeks')
stocklist.pcbt$'Change from Last Four Weeks'<-format(round(stocklist.pcbt$'Change from Last Four Weeks', 2), nsmall = 2)
stocklist.pcbt$'Change from Last Four Weeks'<-as.numeric(stocklist.pcbt$'Change from Last Four Weeks')
stocklist.pcbt$'Change from Last Four Weeks'<-sprintf("%1.2f%%", stocklist.pcbt$'Change from Last Four Weeks')
stocklist.pcbt$'Change from Last 52 Weeks'<-as.numeric(stocklist.pcbt$'Change from Last 52 Weeks')
stocklist.pcbt$'Change from Last 52 Weeks'<-format(round(stocklist.pcbt$'Change from Last 52 Weeks', 2), nsmall = 2)
stocklist.pcbt$'Change from Last 52 Weeks'<-as.numeric(stocklist.pcbt$'Change from Last 52 Weeks')
stocklist.pcbt$'Change from Last 52 Weeks'<-sprintf("%1.2f%%", stocklist.pcbt$'Change from Last 52 Weeks')

stocklist.pcbt$'Latest Quote'<-as.numeric(stocklist.pcbt$'Latest Quote')
stocklist.pcbt$'Latest Quote'<-dollar(stocklist.pcbt$'Latest Quote')

# -------------------------------------------------------------------------


########### Change for Indices Sheet

stocklist.nat$'Change from Previous Week'<-as.numeric(stocklist.nat$'Change from Previous Week')
stocklist.nat$'Change from Previous Week'<-format(round(stocklist.nat$'Change from Previous Week', 2), nsmall = 2)
stocklist.nat$'Change from Previous Week'<-as.numeric(stocklist.nat$'Change from Previous Week')
stocklist.nat$'Change from Previous Week'<-sprintf("%1.2f%%", stocklist.nat$'Change from Previous Week')
stocklist.nat$'Change from Last Four Weeks'<-as.numeric(stocklist.nat$'Change from Last Four Weeks')
stocklist.nat$'Change from Last Four Weeks'<-format(round(stocklist.nat$'Change from Last Four Weeks', 2), nsmall = 2)
stocklist.nat$'Change from Last Four Weeks'<-as.numeric(stocklist.nat$'Change from Last Four Weeks')
stocklist.nat$'Change from Last Four Weeks'<-sprintf("%1.2f%%", stocklist.nat$'Change from Last Four Weeks')
stocklist.nat$'Change from Last 52 Weeks'<-as.numeric(stocklist.nat$'Change from Last 52 Weeks')
stocklist.nat$'Change from Last 52 Weeks'<-format(round(stocklist.nat$'Change from Last 52 Weeks', 2), nsmall = 2)
stocklist.nat$'Change from Last 52 Weeks'<-as.numeric(stocklist.nat$'Change from Last 52 Weeks')
stocklist.nat$'Change from Last 52 Weeks'<-sprintf("%1.2f%%", stocklist.nat$'Change from Last 52 Weeks')

stocklist.nat$'Latest Quote'<-as.numeric(stocklist.nat$'Latest Quote')

######## Write to Excel

write_xlsx(list("Best Performers" = stocklist.top5winners,
                "Worst Performers" = stocklist.top5losers,
                "Pacific Coast Business Times" = stocklist.pcbt,
                "Index" = stocklist.nat),
           path = fname,
           col_names = T)

####################################################
############# Send the email #######################
####################################################

#specify the client secret file. This is necessary to give R
#permission to send from your email address. Link to instructions
#for setting this up: https://www.infoworld.com/article/3398701/how-to-send-email-from-r-and-gmail.html
# gm_auth_configure(path=readLines("../client_secret.json"))
# 
# #text of the message
# message_text <- paste("Hi Amber,\n The stock report for this week is attached.\n Best,\n Sarah\n")
# 
# #more details about the email
# stocks_message <- gm_mime() %>%
#   gm_to("ahair@pacbiztimes.com") %>%
#   gm_from("your_email@ucsb.edu") %>%
#   gm_subject("Stocks for this week") %>%
#   gm_text_body(message_text) %>%
#   gm_attach_file(fname)
# 
# gm_send_message(stocks_message)  
# 1
# 1
# 
# 
# 






