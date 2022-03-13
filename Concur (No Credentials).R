library('httr')
library('jsonlite')
library('lubridate')
library('rvest')
library('jsonlite')
library('xml2')
library('tidyverse')
library('svDialogs')
library('tibble')
library('writexl')
library('aws.s3')
library('pdftools')

####################  GET USERS  ##########################################


#Need to generate token in Postman (See Confluence Article) and enter when prompted
token <- (dlgInput("Copy/Paste Postman Generated Token", '')$res)


#Base URL and initalizing variables used in loop
url <- "https://www.concursolutions.com/api/v3.0/common/users?limit=100&active=TRUE"
run <- 1
row <- 101


#While row is 101. Lol. Because the limit of each API call is 100 rows, it returns what offset value
#is needed to generate the next 100 rows of data in the 101st row, if there is more. So each iteration
#of this loop resets the row variable to the number of rows, and it repeats when that value is 101
#when this loop gets to the end of the data, the row value will be something other than 101 and the
#else statement will be triggered, which ends in a break statement.
while (row == 101){

#Call API
result <- GET(url,
    add_headers( Authorization = paste("Bearer",token, sep = " " )))

#First part of cleaning XML to DataFrame and update row value
list <- as_list(read_xml(result))
tibble <- as_tibble(list)
unnested <- unnest(tibble, cols = Users)
row <- nrow(unnested)

#if row = 101 that means there is another page of data to grab, so the URL is updated and the 101st row trimmed
#after that, the second part of cleaning XML to useable DF occurs
if (row == 101){
  url <- unlist(unnested[101,1])
  trimmed <- head(unnested,-1)
  df <- data.frame(matrix(NA, nrow = 100, ncol = 4))
  
  names(df)[1] <- "Email"
  names(df)[2] <- "User"
  names(df)[3] <- "FirstName"
  names(df)[4] <- "LastName"
  
  y <- 1
  
  for (y in 1:4){
    z <- 1
    for (z in 1:nrow(df)){
      switch (y,
        df[z,y] <- trimmed[[1]][[z]]$LoginID,
        df[z,y] <- trimmed[[1]][[z]]$EmployeeID,
        df[z,y] <- trimmed[[1]][[z]]$FirstName,
        df[z,y] <- trimmed[[1]][[z]]$LastName,
      )
    }
  }
  
  #(re)initalize i, assign output of xml to a numbered dataframe to later be merged using run counter,
  #increase run counter by 1
  i <- 1
  
# for (i in 1:ncol(df)){
#     df[i] <- unlist(df[i])
#   }
  
  assign(paste("output",run,sep=""),df)
  run <- run+1
}

#same as if statement without the updates to url and trimming of 101st row. Ends with break
else{
  vector <- unnest(data.frame(unnested), cols = Users)
  df <- data.frame(matrix(NA, nrow = row, ncol = 4))
  
  names(df)[1] <- "Email"
  names(df)[2] <- "User"
  names(df)[3] <- "FirstName"
  names(df)[4] <- "LastName"
  
  y <- 1
  
  for (y in 1:4){
    z <- 1
    for (z in 1:nrow(df)){
      switch (y,
              df[z,y] <- trimmed[[1]][[z]]$LoginID,
              df[z,y] <- trimmed[[1]][[z]]$EmployeeID,
              df[z,y] <- trimmed[[1]][[z]]$FirstName,
              df[z,y] <- trimmed[[1]][[z]]$LastName,
      )
    }
  }
  
  i <- 1
  
  
  assign(paste("output",run,sep=""),df)
  break
}


}


#Rbind all outputs generated from loop above and remove all concur users to be left with list of active Betenbough users
output.list <- mget(ls(pattern = "output*"))
users <- bind_rows(output.list)
users <- subset(users,!(grepl("Concur", users$Email)))
users <- subset(users,!(grepl("concur", users$Email)))
users <- users[-1,]


################# GET EXPENSE REPORT IDs  ############################

#Create som variable that is mid-month last month to specify date to pull expense reports after
som <- as.Date(format(Sys.Date(), "%Y-%m-01"))
som <- som - 7

#Create IDs which will be a dataframe holder for values generated from following loop.
IDs <- data.frame(c(1:nrow(users)))
names(IDs)[1] <- "ID"
IDs$ID <- "ID"
IDs$User <- "User"
IDs$Total <- "Total"

#For each user, pull a list of expense reports they have submitted after mid-month last month
#and grab the expense report ID of the most recent report that has been submitted for the next looped call
#that will generate each row of each expense report
k <- 1

for (k in 1:nrow(users)){

#API call and cleaning
  result.2 <- GET(paste("https://www.concursolutions.com/api/v3.0/expense/reports?user=",users$Email[k],
                        "&submitDateAfter=",som, sep=""),add_headers( Authorization = paste("Bearer",token, sep = " " )))
  list.2 <- as_list(read_xml(result.2))
  tibble.2 <- as_tibble(list.2)
  unnested.2 <- unnest(tibble.2, cols = Reports)

  
#IF statement handles case when the user has no expense report. If they do, vector2 will have 2 rows and it will extract
#the relevant data. If they don't, it will be length 0 and start at the beginning of the loop again
if (nrow(unnested.2) > 0){
  df.2 <- data.frame(matrix(NA, nrow  = 1, ncol = 3))
  z <- 1
  
  for (z in 1:3){
    switch(z,
           df.2[z] <- unnested.2[[1]][[1]]$ID,
           df.2[z] <- unnested.2[[1]][[1]]$OwnerLoginID,
           df.2[z] <- unnested.2[[1]][[1]]$TotalClaimedAmount
           )
  }
  
  
  IDs$ID[k] <- df.2[1,1]
  IDs$User[k] <- df.2[1,2]
  IDs$Total[k] <- df.2[1,3]
  }
}

#Rows with no report IDs are removed and data in IDs dataframe is unlisted so it is usable
IDs <- subset(IDs,!(grepl("ID",IDs$ID)))



########################### Get Expense Reports ###############################################


#For each expense report ID pulled into IDs dataframe, preform an API call that generates all line items of that report
#rbind all results together to get a dataframe (dataset) that contains all expense report lines to be used for reporting
j <- 1

for (j in 1:nrow(IDs)){
  
#API call
  result.3 <- GET(paste("https://us.api.concursolutions.com/api/expense/expensereport/v2.0/report/",IDs$ID[j],sep=""),
                        add_headers(Authorization = paste("Bearer",token, sep = " ")))

#This response is JSON instead of XML. The below line cleans it into a large list. One element of this large list
#called ExpenseEntriesList is a dataframe with all relevant info in it
  JSON <- fromJSON(rawToChar(result.3$content),flatten = TRUE)
  
#if it is the first run, the dataframe expenses must be initalized. The first line pulls the dataframe from the large list JSON
#the second line removes custom fields that are largely blank from the df
  if (j==1){
    expenses <- data.frame(JSON$ExpenseEntriesList)
    expenses <- expenses[, names(expenses) %in% c("ReportEntryID", "TransactionAmount","CardTransaction.PostedDate",
                                                  "CardTransaction.TransactionDate", "CardTransaction.CardDescription", "VendorDescription")]
  }
  
#After the first run, the data is put into the "holder" dataframe, cleaned of erronious variables, and appended to expense.
  else {
    holder <- data.frame(JSON$ExpenseEntriesList)
    holder <- holder[, names(holder) %in% c("ReportEntryID", "TransactionAmount","CardTransaction.PostedDate",
                                                  "CardTransaction.TransactionDate", "CardTransaction.CardDescription", "VendorDescription")]
    if (ncol(holder) == 6){
    expenses <- rbind(expenses,holder)
  
    }
  }
}


################### GET RECEIPT IMAGES ###########################################

#Initalize loop variable m and create holder df "receipts" to hold URL and ReportEntryID for each image

m <- 1
receipts <- data.frame(c(1:nrow(expenses)))
names(receipts)[1] <- "ID"
receipts$URL <- "URL"


#1:nrow(expenses): For each expense recorded, perform this API call
 for (m in 1:nrow(expenses)){
   image1 <- GET(paste("https://www.concursolutions.com/api/image/v1.0/expenseentry/",expenses$ReportEntryID[m],sep=""),
                 add_headers(Authorization = paste("Bearer", token, sep = " ")))
   
#Convert from JSON to useable dataframe containing ReportEntryID and URL
   image2 <- fromJSON(rawToChar(image1$content),flatten = TRUE)
   image3 <- data.frame(image2)

#If this expense entry has 2 columns, that means there is a receipt associated with the expense
#If that is the case, record the URL and ID into receipts dataframe and download file from URL
#Place the downloaded file into a subfolder of the working directory titled "Receipts"
if (ncol(image3) == 2){
    receipts$URL[m] <- as.character(image3$Url)
    receipts$ID[m] <- as.character(image3$Id)
   download.file(receipts$URL[m],paste("Receipts/",expenses$ReportEntryID[m],sep=""),method="curl")
   try(pdf_convert(paste("Receipts/",expenses$ReportEntryID[m],sep=""), format = "png",
                   filenames = paste("Receipts/",expenses$ReportEntryID[m], sep = "")), silent = TRUE)
   }

 }
 
#Remove any entries into receipts that did not have an associated receipt image
receipts <- subset(receipts, receipts$URL != "URL")

#Set System environment with s3 bucket details 
Sys.setenv(
   "AWS_ACCESS_KEY_ID" = #Put Your Access Key Here,
   "AWS_SECRET_ACCESS_KEY" = #Put Your Secret Access Key Here,
   "AWS_DEFAULT_REGION" = #Put your AWS Region Here
 )

#Initalize loop variable. 
z <- 1

#For z in 1:nrow(receipts): For every expense line that has an associated receipt image, do the following loop 
for (z in 1:nrow(receipts)){
 
#Take file named after ReportEntryID from the "Receipts" subfolder in the working directory and upload it
#to s3 bucket with the same name (unique ReportEntryID)
 put_object(
   file = paste("Receipts/",receipts$ID[z],sep=""),
   object = paste(receipts$ID[z]),
   bucket = #Put your bucket name here
 )
   
}

#https://s3.us-east-1.amazonaws.com/com.betenbough.buckets.financialimages/

##################### OUTPUT ###################################################

colnames(expenses) <- sub("CardTransaction.","",colnames(expenses))
expenses$PostedDate <- as.Date(expenses$PostedDate)
expenses$TransactionDate <- as.Date(expenses$TransactionDate)

#Output results to xlsx with date in file.
write_xlsx(expenses,paste("Concur Expenses ", Sys.Date(), ".xlsx",sep=""))



