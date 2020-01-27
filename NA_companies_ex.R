install.packages("readxl") 
library("readxl")
library(stringr)
data <- read_xlsx('NA_companies_data.xlsx')
dim(data)
str(data)
summary(data)

# Nettoyage nom de colonnes
head(colnames(data)[which(nchar(colnames(data)) > 150)], 1)
colnames(data) <- gsub("\t", " ", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étanger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissemet souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci dindiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprie est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dan chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espage\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne \"Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprise dans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiuez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))
colnames(data) <- gsub("Merci de remplir les cases correspondantes à l'activité de votre entreprise. Si votre entreprise a 2 implantations en Espagne, vous indiquez le chiffre 2 dans la colonne Implantation\" sur la ligne \"Espagne\" . Les 5 colonnes qui suivent vous permettent de qualifier le type d'implantation à l'étranger en indiquant le nombre correspondant dans chaque colonne. Merci d'indiquer dans la colonne \"Ville ou Région\" celle qui correspond à votre implantation dans le pays. Si votre entreprise est implantée dans plusieurs villes du pays, merci de les indiquer les unes à la suite des autres, séparées par une virgule. Merci d'indiquer dans la colonne \"Effectif dans le pays\" le chiffre qui correspond au nombre d'employés de votre entreprisdans les implantations du pays. Si votre établissement souhaite développer 3 implantations en Allemagne, vous indiquez le chiffre 3 dans la colonne \"Souhait de développement des implantations à l'étranger\" sur la ligne \"Allemagne\".     ZONE EUROPE", "", colnames(data))

colnames(data) <- str_trim(colnames(data))


#1-Uniformisez le format du nom de l'entreprise
names <- str_trim(as.matrix(data[,1]), side = "both")
names <- tolower(names)
data[,1] <- names

#2-Remplacez la colonne adresse par 3 colonnes : numéro et le nom de la voie (si contenus), code postale (si contenu), ville (si contenues)
address <- as.matrix(data[,2])
cp <- str_extract(address, pattern = "[0-9]{4-5}")
cp
data$cp <- cp


#3-Verifiez et (si nécessaire) corrigez le bon format des numéros de téléphones (bon format : 10 chars)
phone <- data[,3] %>%
  as.matrix() %>%
  str_replace_all(pattern = "[^0-9]", "") %>%
  str_pad(width = 10, pad = "0")
  #str_replace_all(pattern = "^33", "") %>%
  #substr(start = 1, stop = 10)
phone
data[,3] <- phone

#4-Verifiez et (si nécessaire) corrigez le bon format des emails
#https://regexr.com/
#https://www.regextester.com/
#https://www.regular-expressions.info/

email <- str_extract(as.matrix(data[,4]), pattern = "^([a-z0-9_.-]+)@([\\da-z.-]+)\\.([a-z.]{2,6})$")
email
data[,4] <- email

#5-Verifiez et (si nécessaire) corrigez le bon format des sites internet 
web <- str_extract(as.matrix(data[,5]), pattern = "[-a-zA-Z0-9@:%_+.~#?&//=]{2,256}\\.[a-z]{2,4}")
web <- gsub("http://", "", web)
web <- gsub("https://", "", web)
web
data[,5] <- web

#6-Verifiez et (si nécessaire) corrigez le bon format des numéros SIREN et SIRET (si besoin, vous pouvez vous aider avec https://www.societe.com/)
siren <- str_extract(as.matrix(data[,6]), pattern = "[0-9]{9}")
siren
data[,6] <- siren
siret <- str_extract(as.matrix(data[,7]), pattern = "[0-9]{14}")
siret
data[,7] <- siret

#7-Remplacez les colonnes H-Z par la colonne 'type de filière' : pour chaque entreprise elle contiendra la liste des types de filière renseignés.
col_names <- data[,8:25] %>%
  colnames() %>%
  str_extract(., pattern = "\\[.*\\]")

col_names
names(data)[8:25] <- col_names

apply(data[,8:25], MARGIN = 2, FUN = table)

filieres <- data[,26]

for(i in 8:25){
  data[,i] <- ifelse(tolower(as.matrix(data[,i])) == "oui", names(data)[i], "")
}

#8-Verifiez et (si nécessaire) corrigez le bon format des codes NAF (si besoin, vous pouvez vous aider avec https://www.societe.com/)
NA_companies_data %>%
  select("Code NAF :") %>%
  filter(!grepl("\\d{2}\\.?\\d{2}[A-Z]", `Code NAF :`) & !is.na(`Code NAF :`))

#9-Des colonnes [Pays], gardez seulement les colonnes [IMPLANTATION].
col_start <- grep("\\[Allemagne\\] \\[IMPLANTATION\\]",colnames(NA_companies_data))[1]
col_end <- dim(NA_companies_data)[2]
colnames(NA_companies_data)[col_start:col_end]
columns = which(!grepl("\\[IMPLANTATION\\]", colnames(NA_companies_data)[col_start:col_end])) + col_start -1
columns
colnames(NA_companies_data[,columns])[1]

NA_companies_data <- NA_companies_data %>%
  select(-columns)

#10-Remplacez les colonnes [IMPLANTATION] par une unique colonne 'liste pays' : pour chaque entreprise elle contiendra la liste des pays où l'entreprise est installée.
col_start <- grep("\\[Allemagne\\] \\[IMPLANTATION\\]",colnames(NA_companies_data))[1]
col_end <- dim(NA_companies_data)[2]

NA_companies_data$ListePays = ""
for(col in c(col_start:col_end))
  
{
  col_name = colnames(NA_companies_data)[col]
  first_char = regexpr("\\[", col_name)[1] + 1 
  last_char = regexpr("\\]", col_name)[1] -1
  col_name = substr(col_name, first_char, last_char)
  lines = which(!is.na(NA_companies_data[,col]))
  NA_companies_data[lines, "ListePays"] = paste (NA_companies_data[lines, "ListePays"], col_name)
}

