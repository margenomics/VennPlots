\name{Venn5Plot}
\alias{Venn5Plot}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Venn diagrams for five tables of genes.
}
\description{
Function to make Venn diagrams from five tables containing the columns with array ids or gene symbols, ideally both of them. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained.
}
\usage{
Venn5Plot(listG1, listG2, listG3, listG4, listG5, listNames, filename, data4T = NULL, symbols = TRUE, mkExcel = TRUE, colnmes = c("AffyID", "Symbol"), CatCex = 0.8, CatDist = rep(0.1, 5))
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{listG1}{
First list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listG2}{
Second list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listG3}{
Third list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listG4}{
Fourth list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listG5}{
Fifth list of genes to compare in the Venn diagram. It has to be a data.frame with columns "AffyID" or "Symbol", ideally both of them.
}
  \item{listNames}{
Vector with the names of listG1,listG2,listG3,listG4 and listG5.
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
  \item{colnmes}{
Column names corresponding to the array IDs and gene symbols in the data frame data4T.
}
  \item{CatCex}{
Size of the titles in the Venn diagram
}
  \item{CatDist}{
Distance of the titles with respect to the regions in the Venn diagram.
}
}
\details{
If the option symbols is TRUE a column with the array ids has to be also provided to the function in order to create a .xlsx file. If the option  mkExcel is TRUE a data4T data frame has to be provided and has to have several columns with the name scaled at the beginning of the table. 
}
\value{
Venn diagram plot in .pdf format. If the option mkExcel is TRUE an .xlsx file is created where each sheet is a region of the venn diagram.
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
Magdalena Arnal Segura <marnal@imim.es>
}
\examples{

}
