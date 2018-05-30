\name{Venn4Plot.OS}
\alias{Venn4Plot.OS}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Venn diagrams from four lists of genes.
}
\description{
Function to make Venn diagrams from four vectors of genes. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained with mkExcel in .xlsx format.
}
\usage{
Venn4Plot.OS(listG1, listG2, listG3, listG4, listNames, filename, data4T = NULL, mkExcel = TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{listG1}{
First vector of genes to compare in the Venn diagram.
}
  \item{listG2}{
Second vector of genes to compare in the Venn diagram.
}
  \item{listG3}{
Third vector of genes to compare in the Venn diagram.
}
  \item{listG4}{
Fourth vector of genes to compare in the Venn diagram.
}
  \item{listNames}{
Vector with the names of listG1, listG2, listG3 and listG4.
}
  \item{filename}{
Name of the output files.
}
  \item{data4T}{
Data frame with a column named "Symbol" containing gene symbols in listG1, listG2 and listG3.
}
  \item{mkExcel}{
Logical. If TRUE makes an .xlsx file with the lists of genes corresponding to each region.
}
}
\details{
Simple function to make venn diagrams with the comparisons of three vectors. If mkExcel is TRUE a data4T object has to be provided with a column named "Symbol" and genes corresponding to each region of the Venndiagram will be listed in a excel file.
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
