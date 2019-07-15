\name{Venn2Plot.st}
\alias{Venn2Plot.st}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Venn diagrams for two lists of genes.
}
\description{
Function to make Venn diagrams from two vectors containing gene symbols or gene ids. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained.
}
\usage{
Venn2Plot.st(list1, list2, listNames, filename, Table1, Table2, img.fmt = "pdf", ColName = "GeneSymbol", resDir = NULL)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{list1}{
First vector of genes to compare in the Venn diagram.
}
  \item{list2}{
Second vector of genes to compare in the Venn diagram.
}
  \item{listNames}{
Vector with the names of list1 and list2.
}
  \item{filename}{
Name of the output files.
}
  \item{Table1}{
Table with gene information in list1.
}
  \item{Table2}{
Table with gene information in list2.
}
  \item{img.fmt}{
Default "pdf", if "png" the Venn plot is generated in .png image format.
}
  \item{ColName}{
Name of the column containing gene symbols and common in Table1 and Table2.
}
  \item{resDir}{
Results directory path.
}
}
\details{
Function to make venn diagrams with the comparisons of two vectors and to make a .xlsx file with the common genes. Each sheet contains common genes selected in different tables.
}
\value{
Venn diagram plot in .pdf format and a .xlsx file with common genes.
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
Magdalena Arnal Segura <marnal@imim.es>
}
\examples{

}

