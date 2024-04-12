#-------------------------MRF Script for Google Chromebook TV 2021--------------------------------------------------

# Code has been based off the updated script from Annalect EMEA (Contact: Mila.Myslov@omd.com) 
# which was adapted from the original script from Annalect Germany (Contact: andreas.deneke@annalect.com / sophia.ruenzel@annalect.com)

#--------------------------------------------------------------------------------------------------------------------

#---------------------------------
# PROJECT: Google Chromebook 2021
# MARKET: Sweden
# Collaborators: OMD EMEA & OMD UK
# 
#---------------------------------

#-----------------------------------------------------------------------------------------------
# Step 1 - set up the working directory, load the required libraries and MRF specific functions
#----------------------------------------------------------------------------------------------
# Clear the working environment
rm(list = ls()) 

# Load the required libraries
library(openxlsx)
library(xlsx)
library(DT)
library(foreign)
library(dplyr)
library(tidyr)
library(haven)
library(optimx)
library(sjlabelled) #install.packages("sjlabelled")


# directory on onedrive using the username
username <- Sys.getenv("USERNAME")
wd <- paste0("C:/Users/", username,"/OneWorkplace/OMD-Audience Analytics OMDUK - General/Client Work/Google/Google MRF/2021.06 Google Nordics")
# set working directory
setwd(wd)

# source function
source('./Code/UserDefinedFunctions/MRF_UserDefined_Functions.R')

#-----------------------------------------------------------------------------------------------
# Step 2 - use the "data interface sheet" to import the survey dataset
#-----------------------------------------------------------------------------------------------
# Import the dataset
filename <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", sheetName = "data", 
                      rowIndex = 2:3, colIndex = 2, header = F)

merged_data <- read_sav(paste0("./1. Raw Data/survey_data/", filename[[1]][1]))
dim(merged_data)

prewave_data <- merged_data %>% filter(wave == 1, dCountry == 1)
postwave_data <- merged_data %>% filter(wave == 2, dCountry == 1)

dim(prewave_data)
dim(postwave_data)

# Extract the columns names and description for easy navigation 
questbook_prewave <- CreateQuestBook(prewave_data) # need fixing
questbook_postwave <- CreateQuestBook(postwave_data)

# View the questbook
datatable(questbook_prewave)
datatable(questbook_postwave)

# remove all labels
prewave_df <- remove_all_labels(prewave_data) %>% as_tibble()
postwave_df <- remove_all_labels(postwave_data) %>% as_tibble()

#-------------------------------------------------------------------------------------
# NOTE: Postwave is where we have the TV channels and spot timeslots
# therefore the computation will done using Postwave then I'll join with Prewave
# using respondent ID to avoid messing up with column names required for MRF analysis
#-------------------------------------------------------------------------------------

# Import stations
stations <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                      sheetName = "stations",
                      startRow = 2,
                      cols = 1:2,
                      colNames = T)[, c(1,2)]

num_stations <- nrow(stations)
stations_df <- select(postwave_df, stations$variable) %>% data.frame()

# Import timeslots
timeslots <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                       sheetName = "timeslots",
                       startRow = 2,
                       cols = 1:2,
                       colNames = T)[, c(1,2)]

num_timeslots <- nrow(timeslots)
timeslots_df <- select(postwave_df, timeslots$PostWave_Vars) %>% data.frame()

num_combinations <- num_timeslots * num_stations
stations_big <- stations_df[rep(names(stations_df), times = num_timeslots)] #times= ANZAHL ZEITSCHIENEN
timeslots_big <- timeslots_df[rep(names(timeslots_df), each = num_stations)] #each = ANZAHL SENDER

# import station/timeslot - often
station_timeslot <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                              sheetName = "station-timeslot",
                              startRow = 3,
                              cols = 1:2,
                              colNames = TRUE)[, c(1,2)]

station_timeslot_often <- select(postwave_df, station_timeslot$often_used) %>% data.frame()  
station_timeslot_often[is.na(station_timeslot_often)] <- 0

# import station/timeslot - never
station_timeslot_never <- select(postwave_df, station_timeslot$never_used) %>% data.frame()  
station_timeslot_never[is.na(station_timeslot_never)] <- 0

# import the ad spots run
spots <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                   sheetName = "spots", 
                   startRow = 2, 
                   cols = 3, 
                   colNames = TRUE)[1:80, 3]

# 
# Import the report - i.e. reach at each level of ad contact
report <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                    sheetName = "report", 
                    startRow = 2, 
                    cols = 2, 
                    colNames = TRUE)[, c(1,2)]

report <- round(as.numeric(report[, 2]),2)
report <- ifelse(is.na(report), 0, report)

net_reach <-  round(100 - report[1], 0)

# Import the adblock - i.e. ad reach by station at each timeslot
adblock <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                     sheetName = "adblock", 
                     startRow = 2, 
                     cols = 3, 
                     colNames = TRUE)[1:80, 1:3]

# Import the KPI
kpi_binary <- read.xlsx("./1. Raw Data/DataInterface_SE.xlsx", 
                        sheetName = "kpi",
                        startRow = 2,
                        cols = 1:4,
                        colNames = TRUE)[1:2, 1:4]


#-----------------------------------------------------------------------------------------------
# Step 3 - Data Processing: Viewership Frequency, Combining Frequency Viewing with Stations/Timeslot
#-----------------------------------------------------------------------------------------------
#-----------
# Recoding viewing frequency (timeslot) (see .ppt - slide 16 )
# Q1 (frequency of viewing timeslot in general) is recoded into weights
# daily:	1.0
# 5-6 times a week:	0.75
# 3-4times a week 	0.5
# less often:	0.25
# never:	0
# In this step it is important to make clear definitions of which viewing frequency is transformed into which weight - these weights will be used to calculate the viewing probabilities
#------------
# 
# timeslots_big <- case_when(timeslots_big == 1 ~ 1,
#                           timeslots_big == 2 ~ 0.75,
#                           timeslots_big == 3 ~ 0.5,
#                           timeslots_big == 4 ~ 0.25,
#                           timeslots_big == 5 ~ 0,
#                           TRUE ~ 99)

timeslots_big <- ifelse(timeslots_big == 1, 1,
                        ifelse(timeslots_big == 2, 0.75,
                               ifelse(timeslots_big == 3, 0.5,
                                      ifelse(timeslots_big == 4, 0.25,
                                             ifelse(timeslots_big == 5, 0, 99)))))

# Combining frequency (station/timeslot - often/never)  (see .ppt - slide 17)
# Q2 and Q3 are recoded into one single question:
#   answers to Q2 (watched most often) are recoded into a value of 1
#   answers to Q3 (didn't watch at all) are recoded into a value of 0
#   all others (neither 'most often' nor 'not at all') are recoded into a value of 0.5
# These values will also be used as weights in calculating the viewing probabilities
# Again it is important to consider the implications of choosing the weights


station_timeslot_combination <- ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often==1 & station_timeslot_never==0 ,1,
                                       ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often==0 & station_timeslot_never==1 ,0,
                                              ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often == 0 & station_timeslot_never == 0 , 0.5, 
                                                     ifelse(stations_big == 1 & timeslots_big != 0 & station_timeslot_often == 1 & station_timeslot_never == 1 , 0.5,0))))


# Counting and weighting self-reported viewership (see .ppt slide 19)
#For each station and timeslot frequencies of all answers to Q1 are crosstabbed with the newly recoded Q2/Q3

crosstab_list <- list()

for ( i in 1:num_combinations) {
  crosstab_list[[i]] <- prop.table(table(station_timeslot_combination[,i],timeslots_big[,i]))
}



# The crosstabbed frequencies are weighted and aggregated (see .ppt slide 20)
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


# rewriting this block of code to make computation easier
#-----------------------------------------------------------
# crosstab_list_results <- list()
# 
# 
# for(i in 1:num_combinations){
#   crosstab_list_results[[i]] <- 100*( 
#     (crosstab_list[[i]][3,5]*1.00*1) + (crosstab_list[[i]][2,5]*1.00*0.5) + (crosstab_list[[i]][1,5]*1.00*0) + 
#       (crosstab_list[[i]][3,4]*0.75*1) + (crosstab_list[[i]][2,4]*0.75*0.5) + (crosstab_list[[i]][1,4]*0.75*0) + 
#       (crosstab_list[[i]][3,3]*0.50*1) + (crosstab_list[[i]][2,3]*0.50*0.5) + (crosstab_list[[i]][1,3]*0.50*0) + 
#       (crosstab_list[[i]][3,2]*0.25*1) + (crosstab_list[[i]][2,2]*0.25*0.5) + (crosstab_list[[i]][1,2]*0.25*0) + 
#       (crosstab_list[[i]][3,1]*0.00*1) + (crosstab_list[[i]][2,1]*0.00*0.5) + (crosstab_list[[i]][1,1]*0.00*0))
# }
#-----------------------------------------------------------------------------------------------------------------
crosstab_list_res <- list()

for (i in 1:length(crosstab_list)) {
  
  crosstab_list_res[[i]] <- as.data.frame(crosstab_list[[i]], stringsAsFactors = F) %>%
    mutate(ID = i)
  
}

crosstab_df <- do.call("rbind", crosstab_list_res)
crosstab_df$Var1 <- as.numeric(crosstab_df$Var1)
crosstab_df$Var2 <- as.numeric(crosstab_df$Var2)

str(crosstab_df)

crosstab_agg <- crosstab_df %>% 
  mutate(RowMult = 100*(Var1*Var2*Freq)) %>%
  group_by(ID) %>%
  summarise(Agg = sum(RowMult))

# Transformation of self-reported figures using TV reach data (see .ppt slide 22)
# The values calculated from self-reported viewership are adjusted to peoplemeter TV viewing figures
# self-reported viewership for RTL 7-8 p.m.: 36%
# advertising viewership from TV panel: 2,9%
# adjustment factor:p = 2,9% / 36% = 0,081 (for RTL 7-8 p.m)
# For each person in the data these p-values are then used to calculate individual viewing probabilities

P_value_liste <- list()


for (i in 1:num_combinations){
  P_value_liste[[i]] <-   adblock$adblock_reach[i] / crosstab_agg$Agg[i]
}

P_value_vec <- unlist(P_value_liste)


# Calculation of campaign contacts (see slide 23/24)
# The calculated viewing probabilities are then matched with the campaign flight plan
# All p-values are multiplied with no. of spots placed on the corresponding station in the corresponding time slot
# Results are added up to total campaign contacts

P_value_Matrix <- matrix(P_value_vec, nrow = nrow(postwave_df), ncol = num_combinations, byrow = TRUE)
ots_matrix <- P_value_Matrix*timeslots_big*station_timeslot_combination
spot_matrix <- matrix(t(spots), nrow = nrow(postwave_df), ncol = num_combinations, byrow = TRUE)

contacts_matrix <- ots_matrix*spot_matrix
contacts <- rowSums(contacts_matrix)

postwave_finaldf <- cbind(postwave_df, contacts)

# output contacts to turn them into Frequency
#write.csv(postwave_finaldf$contacts, './3. Processed Data/AdContacts_Survey_v2.csv', row.names = F)

#-----------------------------------------------------------------
# Step 4 - Run the control vs exposed analysis data dump
#-----------------------------------------------------------------
# Control vs Exposed classification (see slide 25)

#Threshold = net reach
rMRF_Threshold <- net_reach/100 

rMRF_TV_data <- postwave_finaldf


# R ranks cases by frequency of contact and creates index variable
rMRF_TV_data$index <- as.numeric(rownames(rMRF_TV_data[rank(rMRF_TV_data$contacts, ties.method="first"),]))

#Determine the case that splits the dataset 
threshold_cutoff <- round(nrow(rMRF_TV_data)*(1-rMRF_Threshold))

rMRF_TV_data$exposed <- NA

#Classify control/exposed

for (i in 1:nrow(rMRF_TV_data)){
  if (rMRF_TV_data$index[i] <= threshold_cutoff) rMRF_TV_data$exposed[i]=0
  else if (rMRF_TV_data$index[i] > threshold_cutoff) rMRF_TV_data$exposed[i]=1
}

#------------------------------------------------------------
# Save postwave df with the contact and exposed identified
#-------------------------------------------------------------
write.csv(rMRF_TV_data, './2. Analysis/Proc_Postwave_SE.csv', row.names = F)


#calculate KPI results for control vs exposed

rMRF_results <- list()

# convert the data in the info.xlsx sheet to strings that define the evaluation criteria of KPIs
for (v1 in 1:nrow(kpi_binary)) {
  kpi_binary$kpi_string[v1] <- paste0(kpi_binary$variable[v1],kpi_binary$operator[v1],kpi_binary$value[v1])
}

rMRF_TV_data$exposed <- factor(rMRF_TV_data$exposed)
unique(rMRF_TV_data$exposed)

# Loop through the dataset and get the results of each of the KPIs agains control vs. exposed
for (i in 1:nrow(kpi_binary)) {
  rMRF_results[[i]] <- rMRF_TV_data %>% 
    group_by(exposed) %>% 
    mutate(v1=n()) %>% 
    filter(eval(parse(text=kpi_binary$kpi_string[i]))) %>% 
    mutate(v2=n(),KPI=(v2/v1)) %>%
    complete(nesting(exposed, KPI))%>% 
    summarise(KPI=ifelse(is.na(mean(KPI)), 0, mean(KPI)))%>%
    select(exposed,KPI)
} 

# Create the output for the data dump
df <- bind_rows(rMRF_results)
df <- data.frame("KPI_Name" = kpi_binary$KPI,
                 "KPI_Variable" = kpi_binary$variable,
                 "Control" = df$KPI[df$exposed==0],
                 "Exposed" = df$KPI[df$exposed==1])

#colnames(df) <- c( "KPI_Name", "KPI_Variable", "control", "exposed")

## opens a prompt to write the file name (MAKE SURE TO INCLUDE .xlsx EXTENTION WHEN WRITING!)
write.xlsx(df, file = "./2. Analysis/Control_Exposed_TV_SE.xlsx")

#-------------------------------------------------------------
# Step 5 - Frequency curve analysis data processing
#--------------------------------------------------------------

# Classification of respondents  Aggregation of calculated contacts into deciles (see .ppt slide 31) 
# All MRF respondents are ranked according to their campaign contacts calculated in step 3
# Based on this ranking respondents are assigned to 10 groups of equal size

postwave_curve <- postwave_finaldf %>%
  arrange(contacts) %>%
  mutate(class = cut(seq_along(contacts), 10, labels = FALSE))


# Creation of deciles (based on reporting) (see .ppt slide 33/34)
# The net reach distribution of the real campaign data is re-distributed into 10 equally sized deciles
# For each aggregated decile average contacts are calculated
contact_distribution <- report

contact_distribution_matrix <- matrix(data = 0, nrow = length(contact_distribution), ncol = 10) 

for (row in 1:nrow (contact_distribution_matrix)){
  for (column in 1:ncol (contact_distribution_matrix)){      
    
    if (row == 1) { 
      if (contact_distribution[row] > 10){ 
        contact_distribution_matrix[row,column] = 10 
        contact_distribution[row] = contact_distribution[row]-10
        
      } else if (contact_distribution[row] > 0) { 
        contact_distribution_matrix[row,column] = contact_distribution[row]
        contact_distribution[row] = 0
        
      } else {contact_distribution_matrix[row,column]=0} 
      
    } else { 
      if (sum(contact_distribution_matrix[1:row,column]) == 10){ 
        contact_distribution_matrix[row,column] = 0 
      } else{ 
        if ( (sum(contact_distribution_matrix[1:row,column])+contact_distribution[row]) > 10 )  {
          input_value= 10-sum(contact_distribution_matrix[1:row,column])
          contact_distribution_matrix[row,column] = input_value
          contact_distribution[row] = contact_distribution[row]-input_value
          
        } else { 
          contact_distribution_matrix[row,column] = contact_distribution[row]
          contact_distribution[row] = 0
          
        }
      }
    }
  }
}

contact_distribution_matrix <- round(contact_distribution_matrix,2)

contact_distribution_matrix <- cbind(contact_distribution_matrix, 0:(nrow(contact_distribution_matrix)-1) )

deciles <- rep(NA,10)
for (i in 1:10){
  deciles[i] <- sum(contact_distribution_matrix[,i]*contact_distribution_matrix[,11])/10
}

#freq(deciles)

# average contacts from the MRF deciles are replaced by those derived from the people meter data (see .ppt slide 36)

postwave_curve$contact_class <- rep(0, times = nrow(postwave_curve))

for (i in 1:10) {
  postwave_curve$contact_class[postwave_curve$class == i] <- deciles[i]
}

#freq(postwave_curve$contact_class)

# KPI regression calculation (see .ppt slide 38)
# For each decile values for the dependent variable - parameters of advertising effectiveness - are counted
# The resulting co-ordinates are
# average contacts per decile derived from people meter data (x)
# percentage of persons with positive response (y)
# A regression analysis is performed on these co-ordinates
# The equation of the regression is set per definition: y = c/1+e^(a-b*x)


kpi_filter <- numeric()
for (i in 1:nrow(kpi_binary)) {
  kpi_filter[i] <- paste0(kpi_binary[i,2], kpi_binary[i,3], kpi_binary[i,4])
}

kpi_list <-list()
for( i in 1:nrow(kpi_binary)){
  kpi_list[[i]] <- postwave_curve %>% 
    group_by(class, contact_class) %>% 
    mutate(v1 = n()) %>% 
    filter(eval(parse(text = kpi_filter[i]))) %>% 
    mutate(v2 = n(), v3 = v2/v1) %>% 
    summarise(kpi = mean(v3)*100) %>% 
    ungroup() %>% 
    select(contact_class, kpi) %>% as.data.frame()
}

names(kpi_list) <- kpi_binary[,1]


## Regression function
MRF_Regression_Calculation <- function(x_data, y_data, accuracy = 17) {
  
  ## Residual sum function
  residual_sum_function <- function(parameterset, x_data, y_data){
    
    ym = parameterset[1]
    a = parameterset[2]
    b = parameterset[3]
    
    return(sum ((y_data-( ym/(1+exp(a-b*x_data))))^2))
  }
  
  ## Start point function
  start_points <- function(goal){ 
    
    start_point_matrix <- expand.grid (ym = seq (0,100, goal),
                                       a = seq (-10,10, goal),
                                       b = seq (0,20, goal))
    return (start_point_matrix)
  }
  
  
  start_point <- start_points(accuracy)
  
  
  cat(nrow(start_point)," Runs --- estimated duration: ", round((nrow(start_point)*0.8),0), " seconds \n" )
  
  bestframe <- data.frame() 
  for (i2 in 1:nrow(start_point)){
    
    pari <- c(start_point$ym[i2],
              start_point$a[i2],
              start_point$b[i2])
    
    optimx_output <- optimx(par=pari, fn = residual_sum_function,
                            lower = c(0, -Inf, 0), upper = c(100, Inf, Inf),
                            control = list(all.methods=TRUE, trace=FALSE),
                            x_data=x_data, y_data=y_data)
    
    names(optimx_output)[1:4] <- c("lm","a","b","sumofsquares")
    
    best_parameter <- (optimx_output[order(optimx_output$sumofsquares),][1:4][1,])
    
    bestframe <- rbind (bestframe,best_parameter)
    
    cat("Run ", i2, " of " ,nrow(start_point)) 
  }
  
  cat ("\n \n done \n") 
  
  return(bestframe[order(bestframe$sumofsquares),][1,])
}  



kpi_data <- kpi_list

MRF_KPI_Performance_Function <- function (kpi_data){
  
  results=list()
  for(i in 1:length(kpi_data)){
    cat("Model ", i, " of ",length(kpi_data),":      ") 
    
    x=kpi_data[[i]][,1] 
    y=kpi_data[[i]][,2]
    
    results[[i]] <- MRF_Regression_Calculation(x_data=x, y_data=y, accuracy = 17) #set "accuracy" to 21 for quick testing, set to 17 for final analysis
  }
  return(results)
}

results <-  MRF_KPI_Performance_Function(kpi_list)

for (i in 1:length(kpi_list)){
  if (i==1){ cat("Kpi","\t","lm","\t","a","\t","b","\n")}
  cat (names(kpi_list[i]), ":",(t(results[[i]])[1:3,1]), "\n")
}



names_list <- names (kpi_list)


# calculate r squared for output template
r_squared_function <- function  (parameters = t(results[[i]])[1:3,1] , data = kpi_list[[i]] ){
  
  ym <- parameters[1]
  a <- parameters[2]
  b <- parameters[3]
  x <- data[,1]
  
  y_estimate <- ym/(1+exp(a-b*x))
  y_actual <- data[,2] 
  
  residual_variation <- (y_actual-y_estimate)^2
  variation_y <- (y_actual-mean(y_actual))^2
  
  r_squared<- 1-(sum(residual_variation)/sum(variation_y))
  
  return (r_squared)
}


#Create Output .xlsx from template (MRF_Vorlage.xls) THIS IS THE REASON FOR XLConnect. Our old database can not process .xslx files.  

workbook <- loadWorkbook("./2. Analysis/MRF_Output_Template_TV_SE.xlsx") #, create = TRUE)
#setStyleAction(workbook, XLC$"STYLE_ACTION.NONE")

writeData(workbook,
          kpi_list[[1]][,1],
          sheet = 1,
          startRow = 12,
          startCol = 1,
          colNames = FALSE)
column = 2

for (i in 1:length(names_list)){
  
  writeData(workbook,
            kpi_list[[i]][,2],
            sheet = 1,
            startRow = 12,
            startCol = column,
            colNames = FALSE)
  
  writeData(workbook,
            r_squared_function(),
            sheet = 1,
            startRow = 6,
            startCol = column,
            colNames = FALSE)
  
  writeData(workbook,
            names (kpi_list)[i],
            sheet = 1,
            startRow = 2,
            startCol = column,
            colNames = FALSE)
  
  writeData(workbook,
            t(results[[i]])[1:3,1],
            sheet = 1,
            startRow = 3,
            startCol = column,
            colNames = FALSE)
  column = column +1
}


saveWorkbook(workbook, "./2. Analysis/TV_Regression_Output_SE.xlsx")

#-------------------------------------------------------
# Significance test of Exposed Vs. Control
#-------------------------------------------------------

#t.test()

summary(rMRF_TV_data$Q8_1)
summary(rMRF_TV_data$Q9_1)



t.test(Q8_1 ~ exposed, data = rMRF_TV_data)


df %>% mutate(Uplift = Exposed-Control)

rMRF_TV_data %>%
  mutate(purint = ifelse(Q9_1 > 2, 0, 1)) %>%
  group_by(exposed) %>%
  summarise(kpi_purint = mean(purint, na.rm = T))
  