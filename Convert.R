## ---------------------------
##
## Script name: Convert.R
##
## Purpose of script: Facturatieservice
##
## Author: Dr. Steven de Jong
##
## Date Created: 10-12-2023
##
## Copyright (c) Steven de Jong, 2023
## 


library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(tidyverse)
library(kableExtra)
library(pagedown)
library(htmltools)
library(writexl)

###############
# SETUP 
###############
Maand <- "December" #Voer maand in
eboutput <- TRUE
factuurnrstart <- 00049

# We need to make a directory
dir.create(Maand, showWarnings = FALSE)

# We need to copy the css there
filecssto <- paste0(Maand,"/my-fonts.css")
filelogoto <- paste0(Maand, "/logo.png")
file.copy("Input/my-fonts.css", filecssto, overwrite = TRUE)
file.copy("logo.png", filelogoto)

logomaand <- paste0("./",Maand,"/logo.png")

####
## STEP 1: Create e-boekhouden import file
####

# Load files, axians is cliënten in axians, ag is agenda uit axians, eb is alle relaties in e-boekhouden, ztn is excelbestand met daarin zorgtrajectnummers, zvc is zorverzekeraarscontracten met percentages
axians <- read_excel("Input/Cliënten 07-12-2023 22 47.xlsx",col_types = c("text","text","text","text","date","numeric","text","text","text","text","text","text","text","text","text","text","text","text"), skip = 3)
ag <- read_excel("Input/Agenda export 20231206.xlsx")
eb <- read_excel("Input/Alle relaties Dialogica Psychologisch en Filosofisch Advies 07-12-2023.xlsx", skip = 9)
ztn <- read_excel("Input/Zorgtrajectnummers.xlsx", col_types = c("text", "text","date","text","text","text","text"))
zvc <- read_excel("Input/Verzekeraars.xlsx", col_types = c("text", "text", "text"))

# Tarieven NZA: https://puc.overheid.nl/nza/doc/PUC_752075_22/1/
nza <- read_excel("Input/Tarieven 2023.xlsx", col_types = c("text", "text", "text", "text")) # Prestatiecodelijst

#Date check. We have a problem with the self-entered dates, as they are the wrong way around
eb <- eb %>% mutate(Geboortedatum = as.Date(Geboortedatum, format="%d-%m-%Y"))
eb <- eb %>% mutate(`Startdatum zorgtraject` = as.Date(`Startdatum zorgtraject`, format="%d-%m-%Y"))

# Prep axians file
axians <- rename(axians, 'Code' = 'Cliëntnummer' )
axians <- axians %>% unite("Naam", c("Achternaam", "Voornaam"), sep = ", ")
axians <- axians %>% unite("Adres", c("Straat", "Huisnummer"), sep = " ")
axians <- rename(axians, 'Telefoon' = `Telefoonnummer (mobiel)` )
axians <- rename(axians, 'Email' = 'E-mail')
axians <- cbind(axians, Uitvoerder='94114926')
axians <- axians %>% left_join(ztn, by='Code')

# Remove clients without Zorgtrajectnummer
axians <- axians %>% filter(!is.na(Zorgtrajectnummer))

# Create converted dataframe
conv <- axians %>% select(Code, Naam, Adres, Postcode, Plaats, Email, Telefoon, Geboortedatum, BSN, Zorgtrajectnummer, Startdatum, Uitvoerder, Uzovi, ZVT, Verwijstype, AGBVW, Profiel)
conv <- rename(conv, `Startdatum zorgtraject` = "Startdatum")
conv <- rename(conv, `UZOVI-code` = "Uzovi")
conv <- rename(conv, `Gekozen zorgvraagtype` = "ZVT")
conv <- rename(conv, `AGB-code verwijzer` = "AGBVW")
conv <- rename(conv, `Gb-GGZ profiel` = "Profiel")

# BSN: Leading zeroes are removed when reading file. Add those if number is less than 9 numbers
Func1 <- function(STR) {
   for (i in seq_along(STR)) {
     if(nchar(STR[i]) < 9) {
       STR[i] <- paste0(strrep("0", 9 - nchar(STR[i])), STR[i])
     }
   }
   return(STR)}
conv$BSN <- Func1(conv$BSN)

# AGB leading zeroes are removed, should be 8 numbers
Func2 <- function(STR) {
  for (i in seq_along(STR)) {
    if(nchar(STR[i]) < 8) {
      STR[i] <- paste0(strrep("0", 8 - nchar(STR[i])), STR[i])
    }
  }
  return(STR)}
conv$`AGB-code verwijzer` <- Func2(conv$`AGB-code verwijzer`)

# UZOVI should be 4 numbers
Func3 <- function(STR) {
  for (i in seq_along(STR)) {
    if(nchar(STR[i]) < 4) {
      STR[i] <- paste0(strrep("0", 4 - nchar(STR[i])), STR[i])
    }
  }
  return(STR)}
conv$`UZOVI-code` <- Func3(conv$`UZOVI-code`)

# If we want to, create an e-boekhouden import file. We should do this now, before renaming variables
if (eboutput == TRUE) {
  ebimport <- conv %>% select(Code,Naam,Adres,Postcode,Plaats,Telefoon,Email,Geboortedatum,BSN,Zorgtrajectnummer,`Startdatum zorgtraject`,`UZOVI-code`,`Gekozen zorgvraagtype`,Verwijstype,`AGB-code verwijzer`,`Gb-GGZ profiel`)
  ebimport <- ebimport %>% mutate(Geboortedatum = as.Date(Geboortedatum, format="%d-%m-%Y"))
  ebimport$Geboortedatum <- format(ebimport$Geboortedatum, "%d-%m-%Y")
  ebimport <- ebimport %>% mutate(`Startdatum zorgtraject` = as.Date(`Startdatum zorgtraject`, format="%d-%m-%Y"))
  ebimport$`Startdatum zorgtraject` <- format(ebimport$`Startdatum zorgtraject`, "%d-%m-%Y")
  write_xlsx(ebimport, paste0(Maand,"/importeb.xlsx"))
  
}

###
# Step 2: Load last month's planning
###

nza <- nza %>% filter(grepl('Ambulant – kwaliteitsstatuut sectie II Gezondheidszorgpsycholoog', Naam)) # Filter GZ-psycholoog

# Komma naar punt
nza$`Maximum-tarief (in Euro)` <- readr::parse_number(nza$`Maximum-tarief (in Euro)`, locale = locale(decimal_mark = ","))


# Filter NA from afspraakcode
ag <- ag %>% filter(!is.na(Afspraakcode))

# We moeten ook van kennismakingsgesprekken af, voor het moment. Die gaan naar een apart bestand
kmg <- "Kennismakingsgesprek"
kmgs <- ag %>% filter(Afspraakcode == kmg)
writexl::write_xlsx (kmgs, paste0(Maand,"/kennismakingsgesprekken.xlsx"))

ag <- ag %>% filter(Afspraakcode != kmg)

# Probably useful to separate client code out of Onderwerp
ag$Code <- str_split_i(ag$Onderwerp, " - ", 1)
# And type
ag$Type <- str_split_i(ag$Afspraakcode, " ", 1)

# a.m. and p.m. to AM and PM
ag$Begin <- str_replace_all(ag$Begin, 'a.m.', 'AM')
ag$Einde <- str_replace_all(ag$Einde, 'a.m.', 'AM')
ag$Begin <- str_replace_all(ag$Begin, 'p.m.', 'PM')
ag$Einde <- str_replace_all(ag$Einde, 'p.m.', 'PM')
ag$Begin <- parse_date_time(ag$Begin, '%m-%d-%Y %I:%M %p')
ag$Einde <- parse_date_time(ag$Einde, '%m-%d-%Y %I:%M %p')

# We hebben intakes als aparte code, maar dat kent de NZA niet
ag$Type <- str_replace_all(ag$Type, "Intake", "Diagnostiek")

# We hebben afspraakdatums nodig in het juiste format, idem voor tijd en duur
ag$afsprdatum <- format(ag$Begin, "%d-%m-%Y")
ag$afsprtijd <- format(ag$Begin,format="%H:%M")
ag$Duur <- difftime(ag$Einde, ag$Begin, units="mins")

# We kunnen nu zorgen dat in de naam netjes de duur komt te staan, met minuten erachter. Zo is de naam identiek aan het bestand van de NZA
ag$Naam <- paste0("Ambulant – kwaliteitsstatuut sectie II Gezondheidszorgpsycholoog (Wet Big artikel 3) ", ag$Type," ", ag$Duur, " minuten")

# De Omschrijving bereiden we vast voor, zodat die op de factuur kan komen. We gebruiken tags uit HTMLtools om ervoor te zorgen dat die straks goed op de factuur kunnen worden geprint
ag$Omschrijving <- paste0("Datum: ", ag$afsprdatum," Tijd: ", ag$afsprtijd, tags$br(), ag$Naam, ". Uitvoerder: 94114926", tags$br(),tags$br())

# Nu kunnen we de agenda joinen op naam
agjoin <- left_join(ag,nza,by="Naam")
agjoinclient <- left_join(agjoin,conv, by="Code")

# Safety check: we moeten van mensen af waar niet alle data is ingevuld, maar die moeten niet gewoon verdwijnen
nietdeclarabel <- agjoinclient %>% filter(is.na(Zorgtrajectnummer))
writexl::write_xlsx (nietdeclarabel, paste0(Maand,"/nietdeclarabel.xlsx"))
agjoinclient <- agjoinclient %>% filter(!is.na(Zorgtrajectnummer)) # nu kunnen we die eruit mikken


# Let's make a list of all the unique codes
unco <- unique(agjoinclient$Code)

#Let's change the stupid NZA name
agjoinclient <- rename(agjoinclient, "Tarief" = `Maximum-tarief (in Euro)`)
agjoinclient <- rename(agjoinclient, "Prestatie" = `Prestatie_code`)
agjoinclient <- rename(agjoinclient, "Omschrijving2" = "Naam.x")
agjoinclient <- rename(agjoinclient, "Datum" = "afsprdatum")

#Geboortedatum in Dutch format
agjoinclient$Geboortedatum <- format(agjoinclient$Geboortedatum, "%d-%m-%Y")

# Let's make an appointment list 

aptlist <- list()

#########################
# UZOVI en VERGOEDINGEN
#########################
agjoinclient <- left_join(agjoinclient, zvc, by="UZOVI-code")
agjoinclient$Percentage <- readr::parse_number(agjoinclient$Percentage, locale = locale(decimal_mark = "."))
agjoinclient$Tarief <- round(agjoinclient$Tarief*(agjoinclient$Percentage/100), digits=2)

for (i in head(unco, n=2)){
  clientcode <- paste0("client.",i) # clientcode zoals client.190469
  
  clientframecommand <- paste0(clientcode,' <- ', 'subset(agjoinclient, Code == ',i,')') # command to make subset dataframe 
    eval(parse(text = clientframecommand))
  
  clientnaam <- eval(parse(text=paste0(clientcode,"[[1,'Naam.y']]")))
#    print(clientnaam)
  clientadres <- eval(parse(text=paste0(clientcode,"[[1,'Adres']]")))
#  print(clientadres)
  clientpostcode <- eval(parse(text=paste0(clientcode,"[[1,'Postcode']]")))
  clientplaats <- eval(parse(text=paste0(clientcode,"[[1,'Plaats']]")))
  clientgd <- eval(parse(text=paste0(clientcode,"[[1,'Geboortedatum']]")))
  clientuzovi <- eval(parse(text=paste0(clientcode,"[[1,'UZOVI-code']]")))
  clientbsn <- eval(parse(text=paste0(clientcode,"[[1,'BSN']]")))
  clientzorgtn <- eval(parse(text=paste0(clientcode,"[[1,'Zorgtrajectnummer']]")))
  clientzorgtndatum <- eval(parse(text=paste0(clientcode,"[[1,'Startdatum zorgtraject']]")))
  clientverwijzer <- eval(parse(text=paste0(clientcode,"[[1,'AGB-code verwijzer']]")))
  clientprofiel <- eval(parse(text=paste0(clientcode,"[[1,'Gb-GGZ profiel']]")))
  clientzvt <- eval(parse(text=paste0(clientcode,"[[1,'Gekozen zorgvraagtype']]")))
  clientafsprtijd <- eval(parse(text=paste0(clientcode,"[[1,'afsprtijd']]")))

  
  factuurdatum <- format(today(), "%d-%m-%Y")
  clientzorgtndatum <- format(clientzorgtndatum, "%d-%m-%Y")
  


  clientinvoiceframename <-  paste0(clientcode, '.invoice')
    # eval(parse(text=clientinvoiceframename))
  
   clientinvoiceframecommand <- paste0(clientinvoiceframename," <- ", clientcode, " %>% select(Prestatie, Omschrijving, `Tarief`)")
    eval(parse(text=clientinvoiceframecommand))

    

    rmarkdown::render(input = "paged.Rmd",
                      params = list(naam = clientnaam,
                      invoiceframe = eval(parse(text=clientinvoiceframename)),
                      adres = clientadres, 
                      postcode = clientpostcode, 
                      plaats = clientplaats,
                      gd = clientgd, 
                      uzovi = clientuzovi,
                      datum = factuurdatum,
                      bsn = clientbsn, 
                      zorgtn = clientzorgtn,
                      zorgtndatum = clientzorgtndatum,
                      verwijzer = clientverwijzer,
                      profiel = clientprofiel,
                      zvt = clientzvt,
                      maand = logomaand,
                      factuurnr = factuurnrstart),
                      output_file = paste0(Maand,"/",factuurnrstart,"-",clientnaam,".html"))
    chrome_print(paste0(Maand,"/",factuurnrstart,"-",clientnaam,".html"), paste0(Maand, "/",factuurnrstart,"-",clientnaam,".pdf"))
    factuurnrstart <- factuurnrstart+1
}
