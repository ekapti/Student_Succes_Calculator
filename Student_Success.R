# Set working directory



# Read xlsx file .
library(readxl)
sd <- read_xlsx('Book2.xlsx')


# Checking the table 
head(sd)

# extract level
sd$level <- substr(sd$paper_code, 3,3)

# extract gpa_value
library(plyr)
sd$gpa <- revalue(sd$grade,
                  c('AE' = 9, 'A+' = 9, 'A' = 8, 'A-' = 7,
                    'B+' = 6, 'B' = 5, 'B-' = 4,
                    'C+' = 3, 'C' = 2, 'C-' = 1,
                    'FT' = 0, 'FE' = 0, 'FD' = 0, 'FC' = 0,
                    'E' =0, 'DS' = 0, 'D' = 0))
sd$gpa <- as.integer(sd$gpa)
detach(package:plyr)
library(dplyr)

# find if any year_at bigger than 4
any(sd$year_at>4)
sd$year_at[which(sd$year_at>4)]


# Group By USE gpa
# Group By year
sd_year_avg_gpa <- sd %>%
  group_by(student_id,year_at) %>%
  summarise( year_at_mean = mean(gpa),
             year_at_max = max(year_at),
             n = n())

# Maximum year in the list
n <- max(sd_year_avg_gpa$year_at_max)
# List of year_ats
year_at_list <- sort(unique(sd_year_avg_gpa$year_at))

# Group By Level 
sd_level_avg_gpa <- sd %>%
  group_by(student_id,level) %>%
  summarise( level_mean = mean(gpa))


# Wide transfrom tables
library(data.table)
sd_year_spread_gpa <- dcast(setDT(sd_year_avg_gpa), student_id ~ year_at, value.var = c('year_at_mean','n'))
library(tidyr)
sd_level_spread_gpa <- spread(sd_level_avg_gpa, 'level', 'level_mean')


detach(package:data.table)

sd_avg_gpa <- merge(sd_level_spread_gpa, sd_year_spread_gpa, by = 'student_id')

#  CSIS

# Replace 'NA' values in n_i with 0

for (i in c(year_at_list)) {
  a <- paste0('sd_avg_gpa$n_',i,'[is.na(sd_avg_gpa$n_',i,')] = 0')
  print(a)
  eval(parse(text = a)) 
}

#get years greater than 4
year_at_list_4plus <- year_at_list[year_at_list>4]


# Calculate cumulative gpa per each year
for (i in c(year_at_list)) {
  a = paste0("sd_avg_gpa$CSIS_cum_",i," <- apply(sd_avg_gpa[,c('year_at_mean_",i,"','n_",i,"')],1,prod, na.rm = TRUE)")
  print(a)
  eval(parse(text = a))
}  

#get years greater than 4
year_at_list_4plus <- year_at_list[year_at_list>4]

# Calculate CSIS_cum_4plus
sd_avg_gpa$CSIS_cum_4plus <- 0

for (i in c(year_at_list_4plus)) {
  a <- paste0("sd_avg_gpa$CSIS_cum_4plus <- sd_avg_gpa$CSIS_cum_4plus + sd_avg_gpa$CSIS_cum_",i)
  print(a)
  eval(parse(text = a))
}


# Calculate n_4plus
sd_avg_gpa$n_4plus <- 0

for (i in c(year_at_list_4plus)) {
    a <- paste0("sd_avg_gpa$n_4plus <- sd_avg_gpa$n_4plus + sd_avg_gpa$n_",i)
    print(a)
    eval(parse(text = a)) 
}

# Calculate CSIS_4plus

sd_avg_gpa$CSIS_4plus = round(sd_avg_gpa$CSIS_cum_4plus/sd_avg_gpa$n_4plus, 2)

sd_avg_gpa$CSIS_4plus[is.nan(sd_avg_gpa$CSIS_4plus)] <- NA


sd_avg_gpa$CSIS_cum_all <- with(sd_avg_gpa,(CSIS_cum_1+ CSIS_cum_2+CSIS_cum_3+CSIS_cum_4 + CSIS_cum_4plus))
sd_avg_gpa$CSIS_n_all <- with(sd_avg_gpa,(n_1 + n_2 + n_3 + n_4 + n_4plus))
sd_avg_gpa$CSIS_cumulative_avg <- round(with(sd_avg_gpa, CSIS_cum_all/CSIS_n_all),2)                                   




#renaming in here check
sd_avg_gpa <- rename(sd_avg_gpa, CSIS_first= 'year_at_mean_1', CSIS_second = 'year_at_mean_2',
                             CSIS_third = 'year_at_mean_3', CSIS_fourth = 'year_at_mean_4')

sd_avg_gpa <- rename(sd_avg_gpa, CSIS_100 = '1', 
                              CSIS_200 = '2', CSIS_300 = '3', CSIS_400 = '4')


# FInal Table
names(sd_avg_gpa)

sd_final <- sd_avg_gpa[,c('student_id','CSIS_100','CSIS_200','CSIS_300','CSIS_400','CSIS_first',
                          'CSIS_second','CSIS_third','CSIS_fourth','CSIS_4plus','CSIS_cumulative_avg')]
# Extract Table
library(WriteXLS)
WriteXLS(sd_final, "Final_Table.xlsx")

