library(docxtractr)
setwd("C:/Tina_local_file/R_on_windows")

#Read the docx and get the table
#If you have multiple tables then you need to modify the tbl_number
example_docx <- read_docx('NCATS 2021-42 DVR.doc')
#NCATS 2021-42 DVR.doc
tbl_extracted <- docx_extract_tbl(example_docx,tbl_number=13)
#extract table


#Remove asterisk
tbl_extracted$Time.Collected <- gsub("\\*","",tbl_extracted$Time.Collected)

#calculate the actual time difference in hours
tdiff_hr <- difftime(strptime(tbl_extracted$Time.Collected,format='%I:%M'),strptime(tbl_extracted$Time.Dosed,format='%I:%M'),units = "hours")

#strptime(tbl_extracted$Time.Collected,format='%I:%M') -> NA for 7 hr time points "2022-03-13 02:40:00"  daylight savings


tdiff_hr <-  as.numeric(tdiff_hr) #difftime to num
t_num_adjusted_logical <- tdiff_hr < 0 #gives positions at negative values
tdiff_hr[t_num_adjusted_logical] <- tdiff_hr[t_num_adjusted_logical] + 12 #adds 12 to negative time values

tbl_extracted$tdiff_hr <- tdiff_hr #assign tdiff_hr to new column


#change sample.time to numeric
tbl_extracted$Sampling.Time..hr. <- as.numeric(tbl_extracted$Sampling.Time..hr.)

#if sampling.time >= 24, adds sampling.time value to tdiff_hr
b <- tbl_extracted$Sampling.Time..hr. >= 24

tbl_extracted[b,"tdiff_hr"] <- tbl_extracted[b,"tdiff_hr"] + tbl_extracted[b,"Sampling.Time..hr."]

#----------- end of data wrangling. Changing everything to correct format that R can deal with
#ends with a column that has the actual time recorded



