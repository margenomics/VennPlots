\name{Venn4Plot.st}
\alias{Venn4Plot.st}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Venn diagrams for three tables of genes.
}
\description{
Function to make Venn diagrams from three vectors containing gene symbols or gene ids. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained.
}
\usage{
Venn4Plot.st(list1, list2, list3, list4, listNames, filename, Table1, Table2, Table3, Table4, ColName = "GeneSymbol", resDir = NULL)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{list1}{
First vector of genes to compare in the Venn diagram.
}
  \item{list2}{
Second vector of genes to compare in the Venn diagram.
}
  \item{list3}{
Third vector of genes to compare in the Venn diagram.
}
  \item{list4}{
Fourth vector of genes to compare in the Venn diagram.
}
  \item{listNames}{
Vector with the names of list1,list2,list3 and list4.
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
  \item{Table3}{
Table with gene information in list3.
}
  \item{Table4}{
Table with gene information in list4.
}
  \item{ColName}{
Name of the column containing gene symbols and common in Table1, Table2, Table3 and Table4.
}
  \item{resDir}{
Results directory path.
}
}
\details{
Function to make venn diagrams with the comparisons of three vectors and to make a .xlsx file with the common genes. Each sheet contains common genes selected in different tables.
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