
#################################################################
# PROJECT: Google Helpfulness UK - TV - December 2020
#################################################################


library(readxl)
library(tidyverse)
library(openxlsx)
library(foreign)
library(haven)
library(dplyr)
library(ggplot2)
library(sjlabelled)
library(weights) # used for the weighting
library(anesrake) # used for the weighting


user <- Sys.getenv("USERNAME")
setwd(paste0("C:/Users/", user, "/OneWorkplace/OMD-Audience Analytics OMDUK - General/Client Work/Google/Google MRF/2022.04 Search MRF/"))
source('3. Code/UserDefinedFunctions/MRF_UserDefined_Functions.R')


prewave_df <- read_sav("2. Raw Data/ORD-705852-B2L2_2022-05-18_FINAL_W1.sav")
postwave_df <- read.csv("2. Raw Data/postwave_df_with_contacts.csv", stringsAsFactors = F)


#-----------------------------------------------------------------------------------------------
# check for id duplicates and remove
#-----------------------------------------------------------------------------------------------
postwave_df$dupe <- duplicated(postwave_df$recruitment_ID)
prewave_df$dupe <- duplicated(prewave_df$recruitment_ID)

prewave_df <- prewave_df[!(prewave_df$dupe == TRUE),]
postwave_df <- postwave_df[!(postwave_df$dupe == TRUE),]

prewave_df <- subset(prewave_df, select=-dupe)
postwave_df <- subset(postwave_df, select=-dupe)



#-----------------------------------------------------------------------------------------------
# recoding Advocacy and favourable kpi
# anything 1 and 2 <- 1
# anything 3 - 6 <- 0
#-----------------------------------------------------------------------------------------------
# postwave_df$Q49_1[postwave_df$Q49_1 == 2] <- 1
# postwave_df$Q49_1[postwave_df$Q49_1 == 3 | 
#                     postwave_df$Q49_1 == 4 |
#                     postwave_df$Q49_1 == 5 |
#                     postwave_df$Q49_1 == 6] <- 0
# 
# prewave_df$Q49_1[prewave_df$Q49_1 == 2] <- 1
# prewave_df$Q49_1[prewave_df$Q49_1 == 3 | 
#                    prewave_df$Q49_1 == 4 |
#                    prewave_df$Q49_1 == 5 |
#                    prewave_df$Q49_1 == 6] <- 0
# 
# postwave_df$Q48_1[postwave_df$Q48_1 == 2] <- 1
# postwave_df$Q48_1[postwave_df$Q48_1 == 3 | 
#                     postwave_df$Q48_1 == 4 |
#                     postwave_df$Q48_1 == 5 |
#                     postwave_df$Q48_1 == 6] <- 0
# 
# prewave_df$Q48_1[prewave_df$Q48_1 == 2] <- 1
# prewave_df$Q48_1[prewave_df$Q48_1 == 3 | 
#                    prewave_df$Q48_1 == 4 |
#                    prewave_df$Q48_1 == 5 |
#                    prewave_df$Q48_1 == 6] <- 0
# 
# prewave_df <- remove_all_labels(prewave_df) %>% as_tibble()







#-----------------------------------------------------------------------------------------------
#  data interface information
#-----------------------------------------------------------------------------------------------

report <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx", 
                    sheet = "report", 
                    startRow = 2, 
                    cols = 1:2, 
                    colNames = TRUE)

kpi <- read.xlsx("./2. Raw Data/TV Data Interface.xlsx",
                     sheet = "kpi",
                     startRow = 2,
                     cols = 1:4,
                     colNames = TRUE)




#-----------------------------------------------------------------------------------------------
# Create informed citizens flag
#-----------------------------------------------------------------------------------------------
#####
old_prewave <- read_sav("C:/Users/rachel.keyte/OneWorkplace/OMD-Audience Analytics OMDUK - General/Client Work/Google/Google MRF/2020.12 Helpfulness UK/2. Raw Data/2009280_Pre Helpfulness SPSS 121520.sav")
#####



# prewave_df$InformedCitizen <- ifelse(((prewave_df$Q6 %in% c(2:4)) & # employment
#                                         (prewave_df$Q4 %in% c(5:7)) & # education 
#                                         (prewave_df$Q7 %in% c(6:14)) & # income
#                                         (prewave_df$Q19 < 3)), 1, 0) # How often do you consume news
# 
# audience_frame <- select(prewave_df, recruitment_ID, InformedCitizen)
# 
# postwave_df <- merge(postwave_df,audience_frame,by="recruitment_ID")
# 
# paste0("Informed citizens: ", round(sum(postwave_df$InformedCitizen) / nrow(postwave_df) * 100,2), "%")




#-----------------------------------------------------------------------------------------------
# TV Viewing Frequency if needed for removing non tv viewers
# Daily: 1 => 1
# 5 -6 times a week: 2 => 0.75
# 3-4 times per week: 3 => 0.5
# less often: 4 => 0.25
# Never: 5 => 0
#-----------------------------------------------------------------------------------------------

# tv_view_vars <- paste0('Q36NEW_', c(1:10,14))
# 
# myrecode <- function(var) {
#   result <- case_when(var == 1 ~ 1,
#                       var == 2 ~ 0.75,
#                       var == 3 ~ 0.5,
#                       var == 4 ~ 0.25,
#                       var == 5 ~ 0)
# }
# 
# analysis_df <- postwave_df %>%
#   select(pid, contacts, Q61_1, Q63_1, Q49_1, tv_view_vars) %>%
#   mutate_at(vars(all_of(tv_view_vars)), myrecode)
# 
# colnames(analysis_df) %in% tv_view_vars
# 
# analysis_df$ViewScore <- rowSums(analysis_df[, 6:16]/length(tv_view_vars))
# 
# nrow(analysis_df[analysis_df$ViewScore == 0,])
# nrow(analysis_df[analysis_df$ViewScore <= 0.25,])
# 
# # analysis_df <- analysis_df %>%
# #   filter(ViewScore != 0)
# analysis_df <- analysis_df %>%
#   filter(ViewScore > 0.25)
# 
# postwave_df <- subset(postwave_df,pid %in% analysis_df$pid)


#-----------------------------------------------------------------------------------------------
# create a prewave frame of just the people in the postwave
# select only the variables we want from the postwave
#-----------------------------------------------------------------------------------------------
prewave_kpi <- kpi[-nrow(kpi),]
#prewave_df <- subset(prewave_df,recruitment_ID %in% postwave_df$recruitment_ID) %>% select(recruitment_ID,kpi$variable,InformedCitizen)
prewave_df <- subset(prewave_df,recruitment_ID %in% postwave_df$recruitment_ID) %>% select(recruitment_ID,prewave_kpi$variable)

# postwave_df <- postwave_df %>%
#   select(recruitment_ID, Q1, qage, contacts, kpi$variable, InformedCitizen) 
postwave_df <- postwave_df %>%
  select(recruitment_ID, S1, S2, contacts, kpi$variable) #S1 m/f, S2 age 



#-----------------------------------------------------------------------------------------------
# split out into audience datadrames 'all' and 'IC'
# calculate control and exposed for each audiences 
#-----------------------------------------------------------------------------------------------

report <- round(as.numeric(report[, 2]),2)
net_reach <-  round(100 - report[1],0)
reach <- net_reach/100  



# all_df <- postwave_df %>%
#   select(recruitment_ID, contacts, InformedCitizen, kpi$variable) %>%
#   arrange(contacts) 
all_df <- postwave_df %>%
  select(recruitment_ID, contacts, kpi$variable) %>%
  arrange(contacts) 
  
bottom <- round((1-reach) * nrow(all_df), 0)
all_df$exposed <- NA
all_df$exposed[1:bottom] <- 0
all_df$exposed <- ifelse(is.na(all_df$exposed), 1, all_df$exposed)

temp <- all_df[all_df$exposed == 0,]
print(paste0("Max contacts for all control: ", max(temp$contacts)))

#write.csv(all_df, "3. Processed Data/tv_all_df.csv", row.names = F)


# ic_df <- postwave_df %>%
#   subset(InformedCitizen == 1) %>%
#   select(pid, contacts, InformedCitizen, kpi$variable) %>%
#   arrange(contacts) 
# 
# bottom <- round((1-reach) * nrow(ic_df), 0)
# ic_df$exposed <- NA
# ic_df$exposed[1:bottom] <- 0
# ic_df$exposed <- ifelse(is.na(ic_df$exposed), 1, ic_df$exposed)
# 
# temp <- ic_df[ic_df$exposed == 0,]
# print(paste0("Max contacts for ic control: ", max(temp$contacts)))


#-----------------------------------------------------------------------------------------------
# Weighting to nat rep
#-----------------------------------------------------------------------------------------------
audience_frame <- select(postwave_df, recruitment_ID, S1, S2)
all_df <- merge(all_df, audience_frame, by = "recruitment_ID")

write.csv(all_df, "all_df_test.csv")

all_df$weight <- 1
all_df$weight <- ifelse(all_df$S1 == 1 & all_df$S2 < 25,  7.19,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 1 & all_df$S2 >= 25 & all_df$S2 < 30,  2.14,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 1 & all_df$S2 >= 30 & all_df$S2 < 35,  2.12,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 1 & all_df$S2 >= 35 & all_df$S2 < 40,  0.97,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 1 & all_df$S2 >= 40 & all_df$S2 < 45,  0.66,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 1 & all_df$S2 >= 45,  0.53,all_df$weight)

all_df$weight <- ifelse(all_df$S1 == 2 & all_df$S2 < 25,  3.5,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 2 & all_df$S2 >= 25 & all_df$S2 < 30,  1.13,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 2 & all_df$S2 >= 30 & all_df$S2 < 35,  0.96,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 2 & all_df$S2 >= 35 & all_df$S2 < 40,  0.67,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 2 & all_df$S2 >= 40 & all_df$S2 < 45,  0.55,all_df$weight)
all_df$weight <- ifelse(all_df$S1 == 2 & all_df$S2 >= 45,  0.57,all_df$weight)

over35_df <- all_df[all_df$S2 >= 35,]

all_df$S1 <- NULL
all_df$S2 <- NULL
over35_df$S1 <- NULL
over35_df$S2 <- NULL

resp_weights <- all_df[,c("recruitment_ID", "weight")]






                 
                 
#-----------------------------------------------------------------------------------------------
# Main MRF Results - t tests for each audience and kpi
#-----------------------------------------------------------------------------------------------

stats_agg <- data.frame()
statsdid_agg <- data.frame()

Audiences <- c("all", "over35")

for (i in Audiences){
  
  loopframe <- eval(parse(text = paste0( i,"_df")))
  loopframe <- unique(loopframe)
  
  for (j in kpi$variable){
    
    test <- t.test(as.formula(paste0(j," ~ exposed")), data = loopframe,conf.level=0.90)
    
    means <- test$estimate
    pvalues <- test$p.value
    stats <- cbind(t(means),pvalues)
    
    colnames(stats) <- c("Control","Exposed","sig")
    
    stats <- as.data.frame(stats)
    stats$Question <- paste0(j)
    stats$Audience <- paste0(i)
    stats$CountControl <- nrow(subset(loopframe, exposed == 0))
    stats$CountExposed <- nrow(subset(loopframe, exposed == 1))
    
    stats_agg <- rbind(stats_agg,stats) 
  }
  
  # The rest of this loop is to get the weighted version of the audiences
  weight_df <- loopframe %>%
    inner_join(select(unique(postwave_df), recruitment_ID, S1, S2), by = 'recruitment_ID') 
  
  weight_df <- weight_df[weight_df$S1 != 3, ]
  weight_df <- unique(weight_df)
  
  weight_df$S2 <- as.factor(weight_df$S2)
  
  Age_Gender <- weight_df %>% 
    group_by(S1, S2, exposed) %>% 
    summarise(n = n())
  
  control_Data <- filter(weight_df, exposed == 0) %>% as.data.frame()
  exposed_Data <- filter(weight_df, exposed == 1) %>% as.data.frame()
  
  # find the needed proportions of age & gender among the exposed group. 
  # This is used to set the target values that the control group will be weighted to
  age_target <- wpct(exposed_Data$S2)
  gender_target <- wpct(exposed_Data$S1)
  
  targets <- list(gender_target, age_target)
  names(targets) <- c("S1", "S2")
  
  # "raking" or otherwise known as RIMWEIGHTING 
  # Apply the targets to the control group and generate weights
  # exposed has no weight (1), control uses the weighting vector from the raking
  raking <- anesrake(targets, control_Data, control_Data$recruitment_ID, cap = 3, choosemethod = "total")
  
  exposed_Data$Weight <- 1
  control_Data$Weight <- raking$weightvec
  
  final_weight_df <- rbind(exposed_Data, control_Data)  
  
  
  for (j in kpi$variable){
    
    test <- t.test(as.formula(paste0(j,"*Weight ~ exposed")), data = final_weight_df,conf.level=0.90)
    
    means <- test$estimate
    pvalues <- test$p.value
    stats <- cbind(t(means),pvalues)
    
    colnames(stats) <- c("Control","Exposed","sig")
    
    stats <- as.data.frame(stats)
    stats$Question <- paste0(j)
    stats$Audience <- paste0(i,"weighted")
    stats$CountControl <- nrow(subset(final_weight_df, exposed == 0))
    stats$CountExposed <- nrow(subset(final_weight_df, exposed == 1))
    
    stats_agg <- rbind(stats_agg,stats) 
  }
}


#-----------------------------------------------------------------------------------------------
# DiD Results - t tests for each audience and kpi
#-----------------------------------------------------------------------------------------------

# for (i in Audiences){
#   
#   loopframe <- eval(parse(text = paste0( i,"_df")))
#   loopframe <- unique(loopframe)
#   
#   pre_ours <- subset(unique(prewave_df),pid %in% loopframe$pid) %>% 
#     inner_join(select(loopframe, pid, exposed)) %>%
#     mutate(Period = 0)
#   
#   postwave_merge <- loopframe %>%
#     select(pid, kpi$variable, InformedCitizen, exposed) %>%
#     mutate(Period = 1)
#   
#   data <- rbind(pre_ours, postwave_merge)
#   data <- unique(data)
#   data$did <- data$Period * data$exposed
#   
#   
#   for (j in kpi$variable){
#     
#     didreg <- lm(as.formula(paste0(j," ~ Period + exposed + did")), data = data)
#     
#     didlift <- summary(didreg)$coefficients["did","Estimate"]
#     sig <- summary(didreg)$coefficients["did","Pr(>|t|)"]
#     period <- summary(didreg)$coefficients["Period","Estimate"]
#     printexposed <- summary(didreg)$coefficients["exposed","Estimate"]
#     controlpre <- summary(didreg)$coefficients["(Intercept)","Estimate"]
#     controlpost <- controlpre + period
#     exposedpre <- controlpre + printexposed
#     exposedexpected <- exposedpre + period
#     exposedpost <- exposedexpected + didlift
#     
#     statsdid <- as.data.frame(cbind(didlift,sig,controlpre,controlpost,exposedpre,exposedpost,exposedexpected))
#     statsdid$kpi <- j
#     statsdid$audience <- i
#     
#     statsdid$CountControl <- nrow(subset(data, exposed == 0))/2
#     statsdid$CountExposed <- nrow(subset(data, exposed == 1))/2
#     
#     statsdid_agg <- rbind(statsdid_agg,statsdid) 
#     
#   }
#   
# }




for (i in Audiences){
  
  loopframe <- eval(parse(text = paste0( i,"_df")))
  loopframe <- unique(loopframe)
  
  pre_ours <- subset(unique(prewave_df),recruitment_ID %in% loopframe$recruitment_ID) %>% 
    inner_join(select(loopframe, recruitment_ID, exposed)) %>%
    mutate(Period = 0)
  
  postwave_merge <- loopframe %>%
    select(recruitment_ID, prewave_kpi$variable, exposed) %>%
    mutate(Period = 1)
  
  data <- rbind(pre_ours, postwave_merge)
  data <- unique(data)
  data$did <- data$Period * data$exposed
  
  data <- merge(data, resp_weights, by = "recruitment_ID", all.x = T)
  
  for (j in prewave_kpi$variable){
    
    data[[j]] <- data[[j]]*data$weight
    
    didreg <- lm(as.formula(paste0(j," ~ Period + exposed + did")), data = data)
    
    didlift <- summary(didreg)$coefficients["did","Estimate"]
    sig <- summary(didreg)$coefficients["did","Pr(>|t|)"]
    period <- summary(didreg)$coefficients["Period","Estimate"]
    printexposed <- summary(didreg)$coefficients["exposed","Estimate"]
    controlpre <- summary(didreg)$coefficients["(Intercept)","Estimate"]
    controlpost <- controlpre + period
    exposedpre <- controlpre + printexposed
    exposedexpected <- exposedpre + period
    exposedpost <- exposedexpected + didlift
    
    statsdid <- as.data.frame(cbind(didlift,sig,controlpre,controlpost,exposedpre,exposedpost,exposedexpected))
    statsdid$kpi <- j
    statsdid$audience <- i
    
    statsdid$CountControl <- nrow(subset(data, exposed == 0))/2
    statsdid$CountExposed <- nrow(subset(data, exposed == 1))/2
    
    statsdid_agg <- rbind(statsdid_agg,statsdid) 
    
  }
  
}

write.csv(stats_agg, '4. Processed Data/TVMRFUpliftResults_over35.csv', row.names = F)
write.csv(statsdid_agg, '4. Processed Data/TVDiDResults_over35.csv', row.names = F)

