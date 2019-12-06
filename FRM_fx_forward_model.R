
#title: "Valuation of Derivatives"
#author: "Chinwude Nwana & David Fakolujo"
#date: "November 8, 2019"
#output: html_document



#Install and load required packages
#```{r packages}
if (Sys.getenv("JAVA_HOME")!="")
  Sys.setenv(JAVA_HOME="")
library(rJava)
#Install
libraries <-  c("tidyverse", "readxl", "ggplot2", "xlsx", "dplyr", "anytime", "reshape2","lubridate")
for(i in libraries) {
  if(i %in% installed.packages() == FALSE) {install.packages(i)}
}
#Load 
lapply(libraries, library, character.only = TRUE)
#```
#Import excel sheet
#```{r}
#DataRequirement_Path <- file.choose()
DataRequirement_Path<-("C:/Users/Training/Documents/inetpub/python/flask/frm_api/assets/inputs/new/Valuation Model Zenith(USDNGN).xlsx")
DataRequirement<-read.xlsx(DataRequirement_Path, sheetName = 1)
#```


#Determine Cap Rate (99% percentile of historical exchange rates)
#```{r Cap rate}
min_cap <- DataRequirement$Historical.Zero.Rate[which.min(DataRequirement$Historical.Zero.Rate[1:500])]

Historical_Rate <- DataRequirement$Historical.Zero.Rate[1:500]
max_cap <- quantile(Historical_Rate[1:(length(Historical_Rate[!is.na(Historical_Rate)]))], 0.99)
#```


#Create interpolation tables for both curves
#```{r interpolation tables}
min <- min(DataRequirement$Local.Date,DataRequirement$Foreign.Date,na.rm = TRUE)
max_local<-max(DataRequirement$Local.Date,na.rm = TRUE)
max_foreign<-max(DataRequirement$Foreign.Date,na.rm = TRUE)
max <- min(max_local,max_foreign)
Date <- seq(from=min, to=max,by = 1)
interpolation_table <- data.frame(Date, Local.Yield=0, Foreign.Yield=0, Days=0)
#```

#defining linear interpolation function 
#```{r define Interpolate function}
#local_curve
lininterpolate_local <- 
  for(i in 1:length(interpolation_table$Date)){
    Low = 1
    High = length(DataRequirement$Local.Date[!is.na(DataRequirement$Local.Date)])
    Med = as.integer((Low + High)/2)
    repeat{
      if(DataRequirement$Local.Date[Med] < as.Date(interpolation_table$Date[i],origin = "1970-01-01")){Low = Med}
      else{High = Med} 
      
      if(abs(High - Low) <= 1){
        break}
      else{Med = as.integer((Low +High)/2)}
    }
    xBelow = DataRequirement$Local.Date[Low]
    xAbove = DataRequirement$Local.Date[High]
    yBelow = DataRequirement$Local.Zero.Rate[Low]
    yAbove = DataRequirement$Local.Zero.Rate[High]
    xVal   = interpolation_table$Date[i]
    
    yVal =  yBelow + as.numeric(xVal - xBelow) * (yAbove - yBelow) / as.numeric(xAbove - xBelow)
    
    interpolation_table$Local.Yield[i]=yVal
  }

#foreign_curve
lininterpolate_foreign <- 
  for(i in 1:length(interpolation_table$Date)){
    Low = 1
    High = length(DataRequirement$Foreign.Date[!is.na(DataRequirement$Foreign.Date)])
    Med = as.integer((Low + High)/2)
    repeat{
      if(DataRequirement$Foreign.Date[Med] < as.Date(interpolation_table$Date[i],origin = "1970-01-01")){Low = Med}
      else{High = Med} 
      
      if(abs(High - Low) <= 1){
        break}
      else{Med = as.integer((Low +High)/2)}
    }
    xBelow = DataRequirement$Foreign.Date[Low]
    xAbove = DataRequirement$Foreign.Date[High]
    yBelow = DataRequirement$Foreign.Zero.Rate[Low]
    yAbove = DataRequirement$Foreign.Zero.Rate[High]
    xVal   = interpolation_table$Date[i]
    
    yVal =  yBelow + as.numeric(xVal - xBelow) * (yAbove - yBelow) / as.numeric(xAbove - xBelow)
    
    interpolation_table$Foreign.Yield[i]=yVal
  }
#``` 

#```{r}
Reporting_date<- DataRequirement$Valuation.Date[1]
for(i in 1:length(interpolation_table$Date)){
  Day_counter <- as.numeric(as.Date(interpolation_table$Date[i],origin = "1970-01-01") - Reporting_date)
  interpolation_table$Days[i]=Day_counter
}
#```

#Determine Interest Rate Parity
#```{r Interest Rate Parity}
Interest_Rate_Parity <- exp((interpolation_table$Local.Yield*interpolation_table$Days/DataRequirement$DayCountConvention..Local.[1])- (interpolation_table$Foreign.Yield*interpolation_table$Days/DataRequirement$DayCountConvention..Foreign.[1]))
#```

#Determine Forward Rate with no Cap
#```{r Forward Rate with no Cap}
number_of_contracts <- length(DataRequirement$S.N[!is.na(DataRequirement$S.N)])
forward_nocap <- DataRequirement$Spot.Fx[1]*Interest_Rate_Parity
forward_rate_nocap <- vector("numeric", number_of_contracts)
for(i in 1:number_of_contracts) {
  forward_rate_nocap[i] <- forward_nocap[which(interpolation_table$Date == DataRequirement$Mat..Date[i])] 
}

#```

#Determine Forward Rate with Cap
#```{r Forward Rate with Cap}
forward_cap <- ifelse(forward_nocap >= max_cap, max_cap, forward_nocap) 
forward_rate_cap <- vector("numeric", number_of_contracts)
for(i in 1:number_of_contracts) {
  forward_rate_cap[i] <- forward_cap[which(interpolation_table$Date == DataRequirement$Mat..Date[i])] 
}

#```

#Determine Discount Factor
#```{r Discount Factor}
DF <- 1/(exp(interpolation_table$Local.Yield*interpolation_table$Days/DataRequirement$DayCountConvention..Local.[1]))
#```

#Determine Discounted Cash Flow with no Cap
#```{r Discounted Cash Flow with no Cap}
CF_Pay <- vector("numeric", number_of_contracts)
for(i in 1:number_of_contracts) {
  CF_Pay[i] <- (DF[which(interpolation_table$Date == DataRequirement$Mat..Date[i])])*(((forward_nocap[which(interpolation_table$Date == DataRequirement$Mat..Date[i])]) - (DataRequirement$Strike.Rate[i]))*DataRequirement$Amount[i])
}
CF_Rec <-  vector("numeric", number_of_contracts)
for(i in 1:number_of_contracts) {
  CF_Rec[i] <- (DF[which(interpolation_table$Date == DataRequirement$Mat..Date[i])])*((DataRequirement$Strike.Rate[i])-(forward_nocap[which(interpolation_table$Date == DataRequirement$Mat..Date[i])]))*DataRequirement$Amount[i]
}
DCF_NoCap <- ifelse(DataRequirement$Pay.Rec == "Pay", CF_Pay, ifelse(DataRequirement$Pay.Rec == "Rec", CF_Rec, 0))
DCF_NoCap <- DCF_NoCap[1:length(DCF_NoCap[!is.na(DCF_NoCap)])]
#```

#Determine Discounted Cash Flow with Cap
#```{r Discounted Cash Flow with Cap}
CF_PayCap <- vector("numeric", number_of_contracts)
for(i in 1:number_of_contracts) {
  CF_PayCap[i] <- (DF[which(interpolation_table$Date == DataRequirement$Mat..Date[i])])*(((forward_cap[which(interpolation_table$Date == DataRequirement$Mat..Date[i])]) - (DataRequirement$Strike.Rate[i]))*DataRequirement$Amount[i])
}

CF_RecCap <-  vector("numeric", number_of_contracts)
for(i in 1:number_of_contracts) {
  CF_RecCap[i] <- (DF[which(interpolation_table$Date == DataRequirement$Mat..Date[i])])*((DataRequirement$Strike.Rate[i])-(forward_cap[which(interpolation_table$Date == DataRequirement$Mat..Date[i])]))*DataRequirement$Amount[i]
}

DCF_Cap <- ifelse(DataRequirement$Pay.Rec == "Pay", CF_PayCap, ifelse(DataRequirement$Pay.Rec == "Rec", CF_RecCap, 0))
DCF_Cap <- DCF_Cap[1:length(DCF_Cap[!is.na(DCF_Cap)])]
#```


#Determines Net Position with no Cap
#```{r Net Position with no Cap}
Net_Position_NoCap <- ifelse(DCF_NoCap < 0, "LIABILITY", "ASSET")
#```

#Determine the Net Position with Cap
#```{r Net Position with Cap}
Net_Position_Cap <- ifelse(DCF_Cap < 0, "LIABILITY", "ASSET")
#```

#Determine results with no Cap
#```{r Results with no Cap}
Client_value <- DataRequirement$Client.s.Valuation[1:length(DataRequirement$Client.s.Valuation[!is.na(DataRequirement$Client.s.Valuation)])]

NoCap_Variance <- DCF_NoCap - Client_value

NoCap_Variance_Percentage <- (NoCap_Variance/Client_value) 

NoCap_Variance_Percentage <- ifelse(Client_value == 0, 0, NoCap_Variance_Percentage)

Results_NoCap <- data.frame(Net_Position_NoCap, forward_rate_nocap, DCF_NoCap, Client_value, NoCap_Variance, NoCap_Variance_Percentage)

Results_names_NoCap <- c("Net Position", "Estimated Forward Rate", "KPMG's Value", "Client's Value", "Variance", "Variance(%)")
names(Results_NoCap) <- Results_names_NoCap
#```

#Determine results with Cap
#```{r Results with Cap}
Cap_Variance <- DCF_Cap - Client_value

Cap_Variance_Percentage <- (Cap_Variance/Client_value)

Cap_Variance_Percentage <- ifelse(Client_value == 0, 0, Cap_Variance_Percentage)


Results_Cap <- data.frame(Net_Position_Cap, forward_rate_cap, DCF_Cap, Client_value, Cap_Variance, Cap_Variance_Percentage)

Results_names_Cap <- c("Net Position", "Estimated Forward Rate", "KPMG's Value", "Client's Value", "Variance", "Variance(%)")
names(Results_Cap) <- Results_names_Cap
#```

#Determine summary with no Cap
#```{r Summary with no Cap}
KPMG_Asset_Valuation_NoCap <- sum(DCF_NoCap[which(Net_Position_NoCap == "ASSET")])


KPMG_Liability_Valuation_NoCap <- sum(DCF_NoCap[which(Net_Position_NoCap == "LIABILITY")])


Client_Asset_Valuation_NoCap <- sum(Client_value[which(Client_value > 0)])


Client_Liability_Valuation_NoCap <- sum(Client_value[which(Client_value < 0)])


NoCap_Asset_Difference <- KPMG_Asset_Valuation_NoCap - Client_Asset_Valuation_NoCap

NoCap_Liability_Difference <- KPMG_Liability_Valuation_NoCap - Client_Liability_Valuation_NoCap

NoCap_Asset_Variance <- (NoCap_Asset_Difference/Client_Asset_Valuation_NoCap)

NoCap_Asset_Variance <- ifelse(Client_Asset_Valuation_NoCap == 0, 0, NoCap_Asset_Variance)

NoCap_Liability_Variance <- (NoCap_Liability_Difference/Client_Liability_Valuation_NoCap)

NoCap_Liability_Variance <- ifelse(Client_Liability_Valuation_NoCap == 0, 0, NoCap_Liability_Variance)

KPMG_Valuation_Difference_NoCap <- KPMG_Asset_Valuation_NoCap + KPMG_Liability_Valuation_NoCap

Client_Valuation_Difference_NoCap <- Client_Asset_Valuation_NoCap + Client_Liability_Valuation_NoCap

NoCap_Difference <- KPMG_Valuation_Difference_NoCap - Client_Valuation_Difference_NoCap

NoCap_Variance_Difference <- (NoCap_Difference/Client_Valuation_Difference_NoCap)

NoCap_Variance_Difference <- ifelse(Client_Valuation_Difference_NoCap == 0, 0, NoCap_Variance_Difference)

Position_NoCap <- c("Asset", "Liability", "")

KPMGs_Valuation_NoCap <- c(KPMG_Asset_Valuation_NoCap, KPMG_Liability_Valuation_NoCap, KPMG_Valuation_Difference_NoCap)

Clients_Valuation_NoCap <- c(Client_Asset_Valuation_NoCap, Client_Liability_Valuation_NoCap, Client_Valuation_Difference_NoCap)

Difference_NoCap <- c(NoCap_Asset_Difference, NoCap_Liability_Difference, NoCap_Difference)

Variance_NoCap <- c(NoCap_Asset_Variance, NoCap_Liability_Variance, NoCap_Variance_Difference)

Summary_NoCap <- data.frame(Position_NoCap, KPMGs_Valuation_NoCap, Clients_Valuation_NoCap, Difference_NoCap, Variance_NoCap)

Summary_names <- c("Position", "KPMG's Valuation", "Client's Value", "Difference", "Variance(%)")
names(Summary_NoCap) <- Summary_names
#```

#Determine summary with Cap
#```{r Summary with Capping}

KPMG_Asset_Valuation_Cap <- sum(DCF_Cap[which(Net_Position_Cap == "ASSET")])

KPMG_Liability_Valuation_Cap <- sum(DCF_Cap[which(Net_Position_Cap == "LIABILITY")])


Client_Asset_Valuation_Cap <- sum(Client_value[which(Client_value > 0)])

Client_Liability_Valuation_Cap <- sum(Client_value[which(Client_value < 0)])


Cap_Asset_Difference <- KPMG_Asset_Valuation_Cap - Client_Asset_Valuation_Cap

Cap_Liability_Difference <- KPMG_Liability_Valuation_Cap - Client_Liability_Valuation_Cap

Cap_Asset_Variance <- (Cap_Asset_Difference/Client_Asset_Valuation_Cap)

Cap_Asset_Variance <- ifelse(Client_Asset_Valuation_Cap == 0, 0, Cap_Asset_Variance)

Cap_Liability_Variance <- (Cap_Liability_Difference/Client_Liability_Valuation_Cap)

Cap_Liability_Variance <- ifelse(Client_Liability_Valuation_Cap == 0, 0, Cap_Liability_Variance)

KPMG_Valuation_Difference_Cap <- KPMG_Asset_Valuation_Cap + KPMG_Liability_Valuation_Cap

Client_Valuation_Difference_Cap <- Client_Asset_Valuation_Cap + Client_Liability_Valuation_Cap

Cap_Difference <- KPMG_Valuation_Difference_Cap - Client_Valuation_Difference_Cap

Cap_Variance_Difference <- (Cap_Difference/Client_Valuation_Difference_Cap)

Cap_Variance_Difference <- ifelse(Client_Valuation_Difference_Cap == 0, 0, Cap_Variance_Difference)

Position_Cap <- c("Asset", "Liability", "")

KPMGs_Valuation_Cap <- c(KPMG_Asset_Valuation_Cap, KPMG_Liability_Valuation_Cap, KPMG_Valuation_Difference_Cap)

Clients_Valuation_Cap <- c(Client_Asset_Valuation_Cap, Client_Liability_Valuation_Cap, Client_Valuation_Difference_Cap)

Difference_Cap <- c(Cap_Asset_Difference, Cap_Liability_Difference, Cap_Difference)

Variance_Cap <- c(Cap_Asset_Variance, Cap_Liability_Variance, Cap_Variance_Difference)

Summary_Cap <- data.frame(Position_Cap, KPMGs_Valuation_Cap, Clients_Valuation_Cap, Difference_Cap, Variance_Cap)

Summary_names_Cap <- c("Position", "KPMG's Valuation", "Client's Value", "Difference", "Variance(%)")
names(Summary_Cap) <- Summary_names_Cap
#```

#Export results and summary
#```{r Export Results and Summary}

wb = loadWorkbook(DataRequirement_Path)

try(removeSheet(wb, sheetName = "Interpolation Results"), silent = TRUE)
try(removeSheet(wb, sheetName = "Results with no Cap"), silent = TRUE)
try(removeSheet(wb, sheetName = "Results with Cap"), silent = TRUE)
try(removeSheet(wb, sheetName = "Summary with no Cap"), silent = TRUE)
try(removeSheet(wb, sheetName = "Summary with Cap"), silent = TRUE)

saveWorkbook(wb, DataRequirement_Path)

interpolation_names <- c("Date", "Local Yield", "Foreign Yield", "Days")
names(interpolation_table) <- interpolation_names
write.xlsx(interpolation_table, 
           file = DataRequirement_Path,
           sheetName = "Interpolation Results",append=TRUE, row.names=FALSE) 

#png("Local Yield.png")
#plot(interpolation_table$Date, interpolation_table$"Local Yield", type = "l", col = "blue", xlab = "Dates", ylab = "Local Yield", main = "Local Rate Chart")
#dev.off()
#png("Foreign Yield.png")
#plot(interpolation_table$Date, interpolation_table$"Foreign Yield", type = "l", col = "blue", xlab = "Dates", ylab = "Foreign Yield", main = "Foreign Rate Chart")
#dev.off()
#wb = loadWorkbook(DataRequirement_Path)
#sheets <- getSheets(wb)[[2]]
#addPicture(file = paste0(getwd(), "/", "Local Yield.png"), sheet = sheets , scale = 1, startRow = 3, startColumn = 6)
#addPicture(file = paste0(getwd(), "/", "Foreign Yield.png"), sheet = sheets , scale = 1, startRow = 28, startColumn = 6)
#saveWorkbook(wb, DataRequirement_Path)

write.xlsx(Results_NoCap, 
           file = DataRequirement_Path,
           sheetName = "Results with no Cap",append=TRUE, row.names=FALSE)

write.xlsx(Results_Cap, 
           file = DataRequirement_Path,
           sheetName = "Results with Cap",append=TRUE, row.names=FALSE)  
write.xlsx(Summary_NoCap, 
           file = DataRequirement_Path,
           sheetName = "Summary with no Cap",append=TRUE, row.names=FALSE)  
write.xlsx(Summary_Cap, 
           file = DataRequirement_Path,
           sheetName = "Summary with Cap",append=TRUE, row.names=FALSE) 

shell.exec(DataRequirement_Path) 
#```