## Read files
library(xlsx)
library(googlesheets)
library(dplyr)
library(tidyverse)
library(googledrive)

## Converting into non-exponential notation
options(scipen = 999)

dir.create("../data")

gs_auth(token = "googlesheets_token.rds")
suppressMessages(gs_auth(token = "googlesheets_token.rds", verbose = FALSE))
gs_user()
gd_token()
drive_auth()
drive_user()

fl_raw <- drive_find(n_max = 10, team_drive = "TEAM OCEAN", pattern = "artikelen_2019.xls")

lst <- list()
for (i in c(1:nrow(fl_raw))) {
  temp <- tempfile(fileext = ".xlsx")
  drive_download(as_id(fl_raw$id[i]), path = temp, overwrite = TRUE)
  lst[i] <- read.xlsx(temp, sheetIndex = 1, colnames = TRUE)
}

df_ba_new <- unnest(lst[[1]])
## eruit halen van een dataframe uit een list loopt nog niet lekker. 


temp <- tempfile(fileext = ".xlsx")
dl <- drive_download(as_id(fl_raw$id[1]), path = temp, overwrite = TRUE)
dl_df <- read.xlsx(temp, sheetIndex = 1, colnames = TRUE)


df_sm <- read.xlsx("../data/Sparta_modellen_2019.xls", sheetIndex = 1, colnames = TRUE)
df_bm <- read.xlsx("../data/Batavus_modellen_2019.xls", sheetIndex = 1, colnames = TRUE)
df_km <- read.xlsx("../data/Koga_modellen_2019.xls", sheetIndex = 1, colnames = TRUE)
df_sa <- read.xlsx("../data/Sparta_artikelen_2019.xls", sheetIndex = 1, colnames = TRUE)
df_ba <- read.xlsx("../data/Batavus_artikelen_2019.xls", sheetIndex = 1, colnames = TRUE)
df_ka <- read.xlsx("../data/Koga_artikelen_2019.xls", sheetIndex = 1, colnames = TRUE)

## dont execute these rules. only used to initialize sheet to find difference in columns. 
## accell_data <- gs_new("mancenteraccell", ws_title = "batavus", input = df_ba[1,], trim = TRUE, verbose = FALSE)


accell_data <- accell_data %>%
  gs_ws_new(ws_title = "sparta", input = df_sa[1,], trim = TRUE, verbose = FALSE) %>%
  gs_ws_new(ws_title = "koga", input = df_ka[1,], trim = TRUE, verbose = FALSE)

## combine first rows in first sheet
## transpose date 
## use logical statement to find where columns deviate
## sparta data frame has usps9-12 columns
## batavus has usps9
## lets remove these columns
dg_sa1 <- subset(df_sa, select = -c(usps9, usps10, usps11, usps12,Frame.artikelhoogte.in.CM))
dg_sm1 <- subset(df_sm, select = -c(usps9, usps10, usps11, usps12,Frame.artikelhoogte.in.CM))
dg_ba1 <- subset(df_ba, select = -c(usps9, Frame.artikelhoogte.in.CM))
dg_bm1 <- subset(df_bm, select = -c(usps9, Frame.artikelhoogte.in.CM))
dg_ka1 <- subset(df_ka, select = -c(Frame.artikelhoogte.in.CM))
dg_km1 <- subset(df_km, select = -c(Frame.artikelhoogte.in.CM))

## Add a brand column and a brandtype column 
dg_sa1 <- df_sa1 %>% mutate(brand = "Sparta", brandtype = "spartaartikelen")
dg_sm1 <- df_sm1 %>% mutate(brand = "Sparta", brandtype = "spartamodellen")
dg_ba1 <- df_ba1 %>% mutate(brand = "Batavus", brandtype = "batavusartikelen")
dg_bm1 <- df_bm1 %>% mutate(brand = "Batavus", brandtype = "batavusmodellen")
dg_ka1 <- df_ka1 %>% mutate(brand = "Koga", brandtype = "kogaartikelen")
dg_km1 <- df_km1 %>% mutate(brand = "Koga", brandtype = "kogamodellen")



## create 1 dataframe with all data. Skip header columns. 
## alldata1 <- rbind(df_sa1, df_sm1,df_km1, df_ka1, df_bm1, df_ba1)
## reduce the number of columns so that Google Sheets can process it
## Check channable to see which columns are mandatory

------------------------------------------------------------------
#create a dataframe for articles only
dg_art <- rbind(dg_sa1, dg_ka1, dg_ba1)

## change class 

#select mandatory columns
dh_art <- dg_art %>% select(PIM.ID, EAN,External.Title, Short.description, Long.description,  
        brand, tl_item.ASSETS.0) 

dh_art$Short.description <- as.character(dh_art$Short.description)
dh_art$Long.description <- as.character(dh_art$Long.description)

## check current NA values in Short description
sum(is.na(dh_art$Short.description))

## replace short description NA with Long description
dh_art <- dh_art %>% mutate(Short.description=replace(Short.description,
          is.na(Short.description), Long.description[is.na(Short.description)]))

# check if short descr contains less NA's
sum(is.na(dh_art$Short.description))

# filter complete cases of mandatory columns
dh_art_full <- dh_art[complete.cases(dh_art), ]

# create dataframe with non-mandatory columns
dh_art_sub <- dg_art %>% select(PIM.ID, Articlenumber, modeljaar, consumentenadviesprijs,
              Model.Number, framemateriaal,Type.frame.DST)

# merge dataframes on PIM.ID
dh_art_tot <- merge(dh_art_full, dh_art_sub, by = "PIM.ID")

# change gender class to character
dh_art_tot$Type.frame.DST <- as.character(dh_art_tot$Type.frame.DST)

# Set correct gender values
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Dames", "Female", dh_art_tot$Type.frame.DST)
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Meisjes", "Female", dh_art_tot$Type.frame.DST)
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Heren", "Male", dh_art_tot$Type.frame.DST)
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Jongens", "Male", dh_art_tot$Type.frame.DST)
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Uni", "Unisex", dh_art_tot$Type.frame.DST)
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Lage instap", "Unisex", dh_art_tot$Type.frame.DST)
dh_art_tot$Type.frame.DST <- ifelse(dh_art_tot$Type.frame.DST == "Extra Lage Instap", "Unisex", dh_art_tot$Type.frame.DST)

# write new csv
write.csv(dh_art_tot, "mancenternew.csv", row.names = FALSE)


-----------------------------------------------------------------------------------------








# #   dg_art_koga <- dg_ka1 %>%
# #   select(PIM.ID, EAN, Articlenumber, External.Title, modeljaar, consumentenadviesprijs, 
# #          Model.Number,Type.frame.DST, framemateriaal, brand, tl_item.ASSETS.0, Short.description) 
# # 
# # dg_art_full_koga <- dg_art_koga[complete.cases(dg_art_koga), ]
# # 
# # dg_art_sub <- dg_art %>%
# #   select(PIM.ID, EAN, Articlenumber, External.Title, modeljaar, consumentenadviesprijs, 
# #   Model.Number,Type.frame.DST, framemateriaal, brand, tl_item.ASSETS.0, Short.description) 
# ##%>%
#   ## mutate(Type.frame.DST=replace(Type.frame.DST, Type.frame.DST = "Damesmono", "Dames"))
# 
# ## complete cases of article  variables
# dg_art_full <- dg_art_sub[complete.cases(dg_art_sub), ]
# 
# # write csv with complete cases
# write.csv(df_art_full, file = "aa_complete.csv", row.names = FALSE)
# 
# ## missing cases
# df_art_miss <- df_art_sub[!complete.cases(df_art_sub),]
# 
# 
# #Create a new worksheet to mancenteraccell
# accell_data <- accell_data %>% gs_ws_new(ws_title = "channablenew", input = newdata, trim = TRUE, verbose = FALSE)
# 
# ## write to csv since google sheets cannot process the amount of data
# olav1 <- write.csv(alldata1, file = "olavaccell1.csv")
# sam1 <- write.csv(alldata1, file = "samaccell.csv")
# 
# ## table ean
# table(newdata$EAN == "NA")
# 
# 
# ## count how many columns are NA in a new dataframe
# na_count <-sapply(alldata, function(y) sum(length(which(is.na(y)))))
# na_count <- data.frame(na_count)