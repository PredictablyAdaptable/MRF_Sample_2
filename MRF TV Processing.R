#########################################################################################################################
### Processing for the data interface file 
### 1) Format spot data for the spots tab 
### 2) Calculate the contact frequencies for the report tab - Need the rest of the data interface to be filled first 
#########################################################################################################################

### Master
rm(list = ls())

library(openxlsx)
library(haven)
library(dplyr)

user <- Sys.getenv("USERNAME")
setwd("C:/Users/andy.seo/OneDrive - OneWorkplace/Desktop/5. Data and Tables/Learning/2022.04 Search MRF/")
source('./3. Code/UserDefinedFunctions/MRF_UserDefined_Functions.R')


prewave_df <- read_sav("2. Raw Data/ORD-705852-B2L2_2022-05-18_FINAL_W1.sav")
postwave_df <- read_sav("2. Raw Data/ORD-705855-Z7M5_2022-05-18_FINAL_W2.sav")

# removing dulicates
postwave_df$dupe <- duplicated(postwave_df$recruitment_ID)
prewave_df$dupe <- duplicated(prewave_df$recruitment_ID)

prewave_df <- prewave_df[!(prewave_df$dupe == TRUE),]
postwave_df <- postwave_df[!(postwave_df$dupe == TRUE),]

prewave_df <- subset(prewave_df, select=-dupe)
postwave_df <- subset(postwave_df, select=-dupe)



questbook_predata <- CreateQuestBook(prewave_df)
questbook_postdata <- CreateQuestBook(postwave_df)





########################################################################################################
# 1) Spot data
########################################################################################################

#-------------------------------------------------------------------------------------
# add timeslot to the spot file using the hour
#-------------------------------------------------------------------------------------

spot_df <- read.xlsx("./1. Background/Helpfulness 2020 Spot List.xlsx")

spot_df <- spot_df[-nrow(spot_df),] ## get rid of row that has total
spot_df$Time <- round(spot_df$Time, 0) ## round time 
spot_df$hour <- NA ## create a column "hour" with NA's in them

for(i in 1:nrow(spot_df)){
  if(nchar(spot_df$Time[i]) == 3){ ### for numeric data in time column that has 3digits
    spot_df$hour[i] <- substr(spot_df$Time[i], start = 1, stop =1) ## save the 1st digit in hr
  }
  if(nchar(spot_df$Time[i]) == 4){ ## for numeric data in time column that has 4 digits 
    spot_df$hour[i] <- substr(spot_df$Time[i], start = 1, stop =2) ## save the 1st and 2nd digit in hr
  }
}


## converting the hour column to a ordinal factor
spot_df$hour[spot_df$hour == "6"] <- "6h-9h"
spot_df$hour[spot_df$hour == "7"] <- "6h-9h"
spot_df$hour[spot_df$hour == "8"] <- "6h-9h"
spot_df$hour[spot_df$hour == "9"] <- "9h-13h"
spot_df$hour[spot_df$hour == "10"] <- "9h-13h"
spot_df$hour[spot_df$hour == "11"] <- "9h-13h"
spot_df$hour[spot_df$hour == "12"] <- "9h-13h"
spot_df$hour[spot_df$hour == "13"] <- "13h-17h"
spot_df$hour[spot_df$hour == "14"] <- "13h-17h"
spot_df$hour[spot_df$hour == "15"] <- "13h-17h"
spot_df$hour[spot_df$hour == "16"] <- "13h-17h"
spot_df$hour[spot_df$hour == "17"] <- "17h-18h"
spot_df$hour[spot_df$hour == "18"] <- "18h-19h"
spot_df$hour[spot_df$hour == "19"] <- "19h-20h"
spot_df$hour[spot_df$hour == "20"] <- "20h-21h"
spot_df$hour[spot_df$hour == "21"] <- "21h-22h"
spot_df$hour[spot_df$hour == "22"] <- "22h-23h"
spot_df$hour[spot_df$hour == "23"] <- "23h-1h"
spot_df$hour[spot_df$hour == "24"] <- "23h-1h"
spot_df$hour[spot_df$hour == "25"] <- "1h-6h"
spot_df$hour[spot_df$hour == "26"] <- "1h-6h"
spot_df$hour[spot_df$hour == "27"] <- "1h-6h"
spot_df$hour[spot_df$hour == "28"] <- "1h-6h"
spot_df$hour[spot_df$hour == "29"] <- "1h-6h"


#-------------------------------------------------------------------------------------
# process the station names 
# create a count of the number of spots in each timeslot and station
#-------------------------------------------------------------------------------------
spot_df$StationNew <- NA

name_map <- read.csv("./1. Background/Station_Name_Mapping.csv", stringsAsFactors = F)

spot_df <- merge(spot_df, name_map, by = "Station" , all.x = T)
spot_df <- spot_df[!is.na(spot_df$StationMap),]


spot_summary <- spot_df %>% 
  group_by(hour, StationMap) %>% 
  tally()

# update the data interface with the spot_summary
write.csv(spot_summary, '/2. Raw Data/Spot_Summary.csv', row.names = F)










########################################################################################################
# 2) Contact frequencies
########################################################################################################



#-------------------------------------------------------------------------------------
# Import information already in the data interface
#-------------------------------------------------------------------------------------

# stations
stations <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx",
                      sheet = "stations",
                      startRow = 2,
                      cols = 1:2,
                      colNames = T)

num_stations <- nrow(stations)
stations_df <- select(postwave_df, stations$variable) %>% data.frame()

# timeslots
timeslots <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx",
                       sheet = "timeslots",
                       startRow = 2,
                       cols = 1:2,
                       colNames = T)

num_timeslots <- nrow(timeslots)
timeslots_df <- select(postwave_df, timeslots$PostWave_Vars) %>% data.frame()

num_combinations <- num_timeslots * num_stations
stations_big <- stations_df[rep(names(stations_df), times = num_timeslots)]                       
timeslots_big <- timeslots_df[rep(names(timeslots_df), each = num_stations)]

# station/timeslot - often watch
station_timeslot <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx",
                              sheet = "station-timeslot",
                              startRow = 3,
                              cols = 1:2,
                              colNames = TRUE)

station_timeslot_often <- select(postwave_df, station_timeslot$often_used) %>% data.frame()
station_timeslot_often[is.na(station_timeslot_often)] <- 0

# station/timeslot - never watch
station_timeslot_never <- select(postwave_df, station_timeslot$never_used) %>% data.frame()
station_timeslot_never[is.na(station_timeslot_never)] <- 0



# spots
spots <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx",
                   sheet = "spots",
                   startRow = 2,
                   cols = 3,
                   colNames = TRUE)


# adblock - i.e. ad reach by station at each timeslot
adblock <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx",
                     sheet = "adblock",
                     startRow = 2,
                     cols = 3,
                     colNames = TRUE)




#-----------------------------------------------------------------------------------------------
# Recoding viewing frequency (timeslot)
# daily:	1.0
# 5-6 times a week:	0.75
# 3-4 times a week 	0.5
# less often:	0.25
# never:	0
#-----------------------------------------------------------------------------------------------

timeslots_big <- ifelse(timeslots_big == 1, 1,
                        ifelse(timeslots_big == 2, 0.75,
                               ifelse(timeslots_big == 3, 0.5,
                                      ifelse(timeslots_big == 4, 0.25,
                                             ifelse(timeslots_big == 5, 0, 99)))))




#-----------------------------------------------------------------------------------------------
# Combining frequency (station/timeslot - often/never)
# answers to watched most often are recoded into a value of 1
# answers to didn't watch at all are recoded into a value of 0
#-----------------------------------------------------------------------------------------------

station_timeslot_combination <- ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often==1 & station_timeslot_never==0 ,1,
                                       ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often==0 & station_timeslot_never==1 ,0,
                                              ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often == 0 & station_timeslot_never == 0 , 0.5, 
                                                     ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often == 1 & station_timeslot_never == 1 , 0.5,0))))



#-----------------------------------------------------------------------------------------------
# Counting and weighting self-reported viewership 
# the two above data frames are crosstabbed
#-----------------------------------------------------------------------------------------------

crosstab_list <- list()

for ( i in 1:num_combinations) {
  crosstab_list[[i]] <- prop.table(table(station_timeslot_combination[,i],timeslots_big[,i]))
}



#-----------------------------------------------------------------------------------------------
# The crosstabbed frequencies are weighted and aggregated
# weights used: 
#   Channels
#     watched most often == 1
#     Neither / nor == 0.5
#     didn't watch at all == 0
#   Timeslots
#     daily == 1
#     5-6 times a week == 0.75
#     3-4 times a week == 0.5
#     less often == 0.25
#     never == 0
#-----------------------------------------------------------------------------------------------

crosstab_list_res <- list()

for (i in 1:length(crosstab_list)) {
  
  crosstab_list_res[[i]] <- as.data.frame(crosstab_list[[i]], stringsAsFactors = F) %>%
    mutate(ID = i)
  
}

crosstab_df <- do.call("rbind", crosstab_list_res)
crosstab_df$Var1 <- as.numeric(crosstab_df$Var1)
crosstab_df$Var2 <- as.numeric(crosstab_df$Var2)

# compute an aggregation - i.e. summing up the row multiplication
crosstab_agg <- crosstab_df %>% 
  mutate(RowMult = 100*(Var1*Var2*Freq)) %>%
  group_by(ID) %>%
  summarise(Agg = sum(RowMult))





#-----------------------------------------------------------------------------------------------
# Transformation of self-reported figures using TV reach data
# The values calculated from self-reported viewership are adjusted using touchpoints data eg
# self-reported value 7-8 p.m.: 36%
# touchpoints: 2,9%
# adjustment factor: p = 2,9% / 36% = 0,081 
# For each person in the data these p-values are then used to calculate individual viewing probabilities
#-----------------------------------------------------------------------------------------------

P_value_liste <- list()


for (i in 1:num_combinations){
  P_value_liste[[i]] <-   adblock[i,] / crosstab_agg$Agg[i]
}

P_value_vec <- unlist(P_value_liste)

P_value_Matrix <- matrix(P_value_vec, nrow = nrow(postwave_df), ncol = num_combinations, byrow = TRUE)


 

#-----------------------------------------------------------------------------------------------
# Calculation of campaign contacts
# ots matrix = All p-values are multiplied with no. of spots placed on the corresponding station in the corresponding time slot
#-----------------------------------------------------------------------------------------------


ots_matrix <- P_value_Matrix*timeslots_big*station_timeslot_combination
spot_matrix <- matrix(t(spots), nrow = nrow(postwave_df), ncol = num_combinations, byrow = TRUE)

contacts_matrix <- ots_matrix*spot_matrix
contacts <- rowSums(contacts_matrix)

postwave_finaldf <- cbind(postwave_df, contacts)
postwave_finaldf$contacts <- round(postwave_finaldf$contacts / 100,0)

write.csv(postwave_finaldf$contacts, './2. Raw Data/AdContacts.csv', row.names = F)
write.csv(postwave_finaldf, './2. Raw Data/postwave_df_with_contacts.csv', row.names = F)


