##authorship listing
#analysts: Kumar Veerapen and Andrea Ganna
#objective: To produce the authorship page for the COVID-19 HGI Flagship paper.
#to note: after producing file, open with MS Word, and replace all ^p with an empty space to get it formatted as shown in paper

##################################
########## Loading file ##########
##################################

# load library for loading file
library("xlsx")


#Kumar's environment for loading author list from paper
#setwd("Dropbox (Partners HealthCare)/HailTeam/COVID19/covid19hgi_authorship/")
#author.list <- read.table("(Final) COVID-19 HGI Authorship and Studies  - Study Authors.tsv", sep="\t", header=T, fill=T, quote = "")
#author.list <- read.xlsx("(Final) COVID-19 HGI Authorship and Studies .xlsx",sheetIndex=1)

#Andrea's environment to load in author list from paper
#author.list <- read.table("/Users/andreaganna/Desktop/(Final) COVID-19 HGI Authorship and Studies  - Study Authors.tsv", sep="\t", header=T, fill=T, quote = "")
#nrow(author.list)
#[1] 3644

author.list <- read.xlsx("/Users/andreaganna/Desktop/(Final) COVID-19 HGI Authorship and Studies .xlsx",sheetIndex=1)
author.list <- author.list[!is.na(author.list$Full.Name),!grepl("NA",colnames(author.list))]

#number of unique authors
length(unique(author.list$Full.Name))
#[1] 3063

#pulling out affiliations, assuming that the split per line is by ;
affiliations <- strsplit(x=as.character(author.list$Affiliation), split=';')
affiliations <- as.data.frame(unlist(affiliations))
nrow(affiliations)
#[1] 3826
#unique affiliations
nrow(unique(affiliations))
#[1] 1180

author.list$Affiliations.split <- strsplit(x=as.character(author.list$Affiliation), split=';')
author.list$Affiliations.split <- sapply(author.list$Affiliations.split, function(x){trimws(x)})


##################################
########## The Script ############
##################################

library(officer)
library(magrittr)
library(dplyr)
library(data.table)

## some fixes
author.list$Role <- trimws(author.list$Role)
author.list$Full.Name <- trimws(author.list$Full.Name)

## Sort the author list
role_ordered <- ordered(author.list$Role,levels=
                          c("Leadership","Writing group lead",
                            
                            "Manuscript analyses team lead","Manuscript analyses team members: phewas","Manuscript analyses team members: mendelian randomization","Manuscript analyses team members: methods development","Manuscript analyses team members: PC projection, gene prioritization",
                            
                            "Scientific communication lead","Website development",
                            
                            "Project management lead","Project managment support","Phenotype steering group","Data dictionary",
                            "Analysis Team Lead","Data Collection Lead","Admin Team Lead","Analysis Team Member","Data Collection Member","Admin Team Member"))

main_groups <- c("Leadership","Writing group","Analysis group","Project management group","Website development","Scientific communication group")
study_ordered <- ordered(author.list$Study,levels=c(main_groups,sort(unique(author.list$Study[!author.list$Study %in% c(main_groups,"COVID-19 HGI corresponding authors")])),"COVID-19 HGI corresponding authors"))

author.list.sorted <- author.list[order(study_ordered,role_ordered),]


### Form affiliation ##  
affiliations <- as.data.frame(do.call(rbind, lapply(author.list.sorted$Affiliations.split, as.data.frame)))
colnames(affiliations) <- c("name")
affiliations$number <- "test"



affiliationsNext <- NULL

#removing duplicate affiliations and ensuring that they are all in the the order for the authors because I am paranoid
for(i in 1:nrow(affiliations)){
  if(is.null(affiliationsNext) == TRUE |
    !(affiliations[i,]$name %in% affiliationsNext$name) == TRUE ) {
    affiliationsNext <- as.data.frame(rbind(affiliationsNext, as.character(affiliations[i,]$name) ),stringsAsFactors = FALSE)
    colnames(affiliationsNext) <- c("name")
    } 
}

affiliations <- affiliationsNext
rm(affiliationsNext)
affiliations$number <- "test"
colnames(affiliations) <- c("name", "number")
head(affiliations)

#adding values to the affiliations for the superscripts
for(i in 1:nrow(affiliations)){
  affiliations[i,]$number <- i
}

#open doc
x <- read_docx()

#Creating formatting properties
studyNameProp <- fp_text(color="black", bold=TRUE, font.size = 14)
roleProp <- fp_text(color="black", bold=TRUE)
nameProp <- fp_text(color="black")
affiliationProp <- fp_text(color="black", vertical.align="superscript")
affiliationNameProp <- fp_text(color="black")
commaSep <- ftext(",", nameProp)
  
for (study in unique(author.list.sorted$Study))
{
  # Select only the study of interest
  author.list.sorted.study <- author.list.sorted[author.list.sorted$Study==study,]
  
  # Add Study name
  studyName <- ftext(study, studyNameProp)
  paragraph <- fpar(run_linebreak(), studyName)
  x <- body_add_fpar(x, paragraph)
  
  for (role in unique(author.list.sorted.study$Role))
  {
    author.list.sorted.study.role <- author.list.sorted.study[author.list.sorted.study$Role==role,]
    
    # Add role name
    roleName <- ftext(role, roleProp)
    paragraph <- fpar(run_linebreak(),roleName)
    x <- body_add_fpar(x, paragraph)
    
    for (i in 1:nrow(author.list.sorted.study.role))
    {
      Name <- ftext(as.character(author.list.sorted.study.role[i,]$Full.Name), nameProp)
      
      # Find all affiliations corresponding to the name
      all_affiliation <- unlist(author.list.sorted.study.role$Affiliations.split[i])
      affiliation_number <- paste(affiliations$number[affiliations$name %in% all_affiliation],collapse=",")
      affiliationSuper <- ftext(affiliation_number, affiliationProp) 
      
      paragraph <- fpar(Name, affiliationSuper, fp_p = fp_par(keep_with_next = TRUE))
      x <- body_add_fpar(x, paragraph, pos="after")
      x <- slip_in_text(x, ",", pos = "after")
    }
  }
}


print(x, target = "/Users/andreaganna/Desktop/covid19hgi_AuthorList.docx")


## Affiliations

x3<- read_docx()

for (k in 1:nrow(affiliations))
{
  textind <- paste0(affiliations$number[k],". ",affiliations$name[k])
  affili <- ftext(textind, nameProp)
  paragraph <- fpar(affili)
  x3 <- body_add_fpar(x3, paragraph)
  print(k)
}

print(x3, target = "/Users/andreaganna/Desktop/covid19hgi_AuthorList_affiliations.docx")


  