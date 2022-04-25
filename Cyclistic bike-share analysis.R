
library(openxlsx)
library(pivottabler)
library(viridisLite)
library(viridis)
library(ggplot2)
library(dplyr)


# Loading all the excel files of last 12 months and apply them

path <- "C:/Users/mehr/Desktop/case study 1/24 months/previous-12-months/excel-file/raw-excel"

merge_file_name <- "C:/Users/mehr/Desktop/case study 1/24 months/previous-12-months/excel-file/raw-excel//merged_file.xlsx"

filenames_list <- list.files(path= path, full.names=TRUE)

All <- lapply(filenames_list,function(filename){
  print(paste("Merging",filename,sep = " "))
  read.xlsx(filename)
})

df <- do.call(rbind.data.frame, All)
#write.xlsx(df,merge_file_name)

################################################################################################

# Data frame summary
summary(df)

str(df)

ls(df)

head(df)


# Data frame dimensions
dim(df)





#Frequency table for categorical values 

table(df$member_casual, useNA = "always")
table(df$day_of_week, useNA = "always")
table(df$month_num, useNA = "always")
table(df$month_of_year, useNA = "always")
table(df$rideable_type, useNA = "always")




#find the missing values in the entire dataframe

sum(is.na(df$member_casual))
sum(is.na(df$day_of_week))
sum(is.na(df$month_num))
sum(is.na(df$month_of_year))
sum(is.na(df$rideable_type)) 
sum(is.na(df$start_station_name))    #has missing value
sum(is.na(df$start_station_id))      #has missing value
sum(is.na(df$end_station_name))      #has missing value
sum(is.na(df$end_station_id))        #has missing value
sum(is.na(df$started_at))
sum(is.na(df$ended_at))
sum(is.na(df$start_lat))
sum(is.na(df$start_lng))
sum(is.na(df$end_lat))               #has missing value
sum(is.na(df$end_lng))               #has missing value



#Finding duplicates

sum(duplicated(df))                  # there is no duplicated data in dataframe


###################################################################################################
# Define function which change float to hours and minute

convertTime <- function(x){
  Hour = (x * 24) # %% 24#For x>1, only the value after decimal will be considered
  Minutes = (Hour %% 1) * 60
  Seconds = (Minutes %% 1) * 60
  
  hrs = ifelse (Hour < 10, paste("0",floor(Hour),sep = ""), as.character(floor(Hour)))
  mins = ifelse (Minutes < 10, paste("0",floor(Minutes),sep = ""), as.character(floor(Minutes)))    
  secs = ifelse (Seconds < 10, paste("0",round(Seconds,0),sep = ""), as.character(round(Seconds,0)))
  
  return(paste(hrs, mins, secs, sep = ":"))
}




#Calculate the max and average ride_length

mean_ride_length_year <-convertTime( mean(df$ride_length))
mean_ride_length_year
max_ride_length_year <-convertTime( max(df$ride_length))
max_ride_length_year


############################################ Data Analysis ########################################

# Calculate the average length of ride  for user type.

pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="mean_ride_length", 
                     summariseExpression="mean(ride_length, na.rm=TRUE)")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("member_casual")
pt$evaluatePivot()
df1 <- pt$asDataFrame()
df1
df1$mean_ride_length_time <-convertTime(df1$mean_ride_length)
df1


# Calculate the count length of ride  for user type.
#Number of rides completed by user type between April 2021 until March 2022

pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="count_ride_length", 
                     summariseExpression="n( )")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("member_casual")
pt$evaluatePivot()
pt$renderPivot()




ggplot(data=df)+geom_bar(aes(x=member_casual))

ggplot(df, aes(x=member_casual)) +
  geom_bar(fill = "gold") +
  labs(
    title = "Number of rides completed by user type",
    subtitle = "For the period between April 2021 until March 2022",
    x = "User type",
    y = "Number of rides (in millions)") +
  geom_text(stat='count', aes(label=..count..), vjust=+2, color="black")






# Calculate the number of rides for users by day_of_week by adding Count of trip_id to Values

pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="count_ride_id", 
                     summariseExpression="n( )")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("day_of_week")
pt$renderPivot()



ggplot(df, aes(x=day_of_week)) +
  geom_bar(fill = "gold") +
  labs(
    title = "Number of rides completed by day of week",
    subtitle = "For the period between April 2021 until March 2022",
    x = "day of week",
    y = "Number of rides (in millions)") +
  geom_text(stat='count', aes(label=..count..), vjust=+2, color="black")





# Calculate the average ride_length for users by day_of_week and user type.

pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation( calculationName=" ",
                     summariseExpression="mean(ride_length, na.rm=TRUE)")
pt$addColumnCalculationGroups()
pt$addColumnDataGroups("day_of_week")
pt$addRowDataGroups("member_casual")
pt$evaluatePivot()
df2 <- pt$asDataFrame()
df2
df2[] <- lapply(df2, convertTime)
df2


df%>%group_by(member_casual,day_of_week)%>%summarise(average_ride_length=mean(ride_length))%>%
  ggplot(aes(x=day_of_week,y=average_ride_length,fill=member_casual)) +
  geom_bar(position="Dodge",stat = "identity") +
  scale_fill_manual(values=c("slateblue1","gold"))+
  labs(title="distribution of average ride length by week and user type")





#Calculate the number of rides for users by member_casual and day_of_week by adding Count of trip_id to Values

pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="count_ride_id", 
                     summariseExpression="n( )")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("member_casual")
pt$addRowDataGroups("day_of_week")
pt$renderPivot()



df%>%group_by(member_casual,day_of_week)%>%summarise(n=n())%>%
  ggplot(aes(x=day_of_week,y=n,fill=member_casual)) +
  geom_bar(position="Dodge",stat = "identity") + 
  scale_fill_manual(values=c("slateblue1","gold"))+
  labs(title="distribution of count ride length by week and user type")
  



#Calculate the number of rides for users by month_of_year and day_of_week and member_casual by adding Count of trip_id to Values


pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="count_ride_id", 
                     summariseExpression="n( )")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("month_of_year")
pt$addRowDataGroups("day_of_week")
pt$addRowDataGroups("member_casual")
pt$renderPivot()



df%>%group_by(day_of_week,member_casual,month_of_year)%>%summarise(n=n())%>%
  ggplot(aes(x=day_of_week,y=n,fill=month_of_year)) +
  geom_bar(position="Dodge",stat = "identity") +
  scale_fill_viridis_d()+
  facet_wrap(~member_casual)+
  labs(title="distribution of count ride length by month and week and user type")






#Calculate the number of rides for users by user type and rideable_type by adding Count of trip_id to Values


pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="count_ride_id", 
                     summariseExpression="n( )")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("member_casual")
pt$addRowDataGroups("rideable_type")
pt$renderPivot()



df%>%group_by(member_casual, rideable_type)%>%summarise(n=n())%>%
ggplot( aes(x=member_casual, y= n, fill=rideable_type)) +
  geom_bar(stat="identity") +
  scale_fill_manual(values=c("slateblue1","seagreen1","gold"))+
  labs(
    title = "Bike preference by user type and rideable type ",
    subtitle = "For the period between April 2021 until March 2022",
    fill = "Bike type") +
  geom_text(aes(label=n), position = position_stack(vjust = .5), color="black") 
 





#Calculate the number of rides for users by month_of_year and user type by adding Count of trip_id to Values

pt <- PivotTable$new()
pt$addData(df)
pt$defineCalculation(calculationName="count_ride_id", 
                     summariseExpression="n( )")
pt$addColumnCalculationGroups()
pt$addRowDataGroups("month_of_year")
pt$addRowDataGroups("member_casual")
pt$renderPivot()


df%>%group_by(month_of_year,member_casual)%>%summarise(n=n())%>%
  ggplot(aes(x=month_of_year,y=n,fill=month_of_year)) +
  geom_bar(position="Dodge",stat = "identity") +
  scale_fill_viridis_d()+
  facet_wrap(~member_casual)+
  labs(title="distribution of count ride length by month_of_year and user type")
