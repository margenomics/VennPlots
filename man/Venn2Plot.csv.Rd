\name{Venn2Plot.csv}
\alias{Venn2Plot.csv}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Venn diagrams for two lists of genes.
}
\description{
Function to make Venn diagrams from two tables containing the columns with array ids or gene symbols, ideally both of them. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained with mkExcel (Excel format) and mkCSV (csv format) options.
}
\usage{
Venn2Plot.csv(listG1, listG2, listNames, filename, data4T = NULL, symbols = TRUE, mkExcel = TRUE, mkCSV = FALSE, colnmes = c("AffyID", "Symbol"))
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{listG1}{
First list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listG2}{
Second list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listNames}{
Vector with the names of listG1 and listG2.
}
  \item{filename}{
Name of the output files.
}
  \item{data4T}{
Table with the structure of data4Tyers with columns named scaled at the begining.
}
  \item{symbols}{
Logical. Use symbols instead of AffyID to make plots.
}
  \item{mkExcel}{
Logical. If TRUE makes an .xlsx file with the lists of genes corresponding to each region.
}
  \item{mkCSV}{
Logical. If TRUE makes a .csv file for the genes corresponding to each region.
}
  \item{colnmes}{
Column names corresponding to the array IDs and gene symbols in the data frame data4T.
}
}
\details{
If the option symbols is TRUE a column with the array ids has to be also provided to the function lists in order to create the .csv and .xlsx files and the data4T data frame has to have several columns with the name scaled at the beginning of the table. 
}
\value{
Venn diagram plot in .pdf format. If the option mkExcel is TRUE an .xlsx file is created where each sheet is a region of the venn diagram. If the option mkCSV is TRUE several .csv files are created corresponding to each region of the venn diagram.
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
Magdalena Arnal Segura <marnal@imim.es>
}
\examples{

}

