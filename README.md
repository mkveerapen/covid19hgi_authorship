# COVID-19 HGI Authorship Listing Page 

## This script was used to allow the creation of the authorship page on the flagship paper

This script was written in `R` and was based on the authorship list that can be found [here](https://docs.google.com/spreadsheets/d/1cp9pFeFUxXz5WMjRFv4X-AM1Hlc0iXYJa1rorSSj2Dc/edit#gid=0).

Thanks to Amy Trankiem and Rachel Liao for curating the list, the main analysts on this are : [Kumar Veerapen](mailto:veerapen@broadinstitute.org) and Andrea Ganna.

## How do you run the code?

Line by line in your terminal. 

**input file** : can be found in the `src/` dir in the spreadsheet format

*to note*: 
1) Author names must not have commas.
2) Multiple affiliations must be split by semicolons. 

**script to run**: can be found in `script/authorshipListing.R`
_main R package used_ : `officeR` (link found [here](https://davidgohel.github.io/officer/))

Using said script, you will be able to obtain a file that looks like: 
`results/covid19hgi_AuthorList_final.docx`

We then fused the lines using MS Word by finding and replacing `^p` with an empty space (` `) which can be seen in `results/covid19hgi_AuthorList_final_KV.docx`
