\name{Venn4Plot.st}
\alias{Venn4Plot.st}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
%%  ~~function to do ... ~~
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
Venn4Plot.st(list1, list2, list3, list4, listNames, filename, Table1, Table2, Table3, Table4, ColName = "GeneSymbol", resDir = NULL)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{list1}{
%%     ~~Describe \code{list1} here~~
}
  \item{list2}{
%%     ~~Describe \code{list2} here~~
}
  \item{list3}{
%%     ~~Describe \code{list3} here~~
}
  \item{list4}{
%%     ~~Describe \code{list4} here~~
}
  \item{listNames}{
%%     ~~Describe \code{listNames} here~~
}
  \item{filename}{
%%     ~~Describe \code{filename} here~~
}
  \item{Table1}{
%%     ~~Describe \code{Table1} here~~
}
  \item{Table2}{
%%     ~~Describe \code{Table2} here~~
}
  \item{Table3}{
%%     ~~Describe \code{Table3} here~~
}
  \item{Table4}{
%%     ~~Describe \code{Table4} here~~
}
  \item{ColName}{
%%     ~~Describe \code{ColName} here~~
}
  \item{resDir}{
%%     ~~Describe \code{resDir} here~~
}
}
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
%%  ~~who you are~~
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
##---- Should be DIRECTLY executable !! ----
##-- ==>  Define data, use random,
##--	or do  help(data=index)  for the standard data sets.

## The function is currently defined as
function (list1, list2, list3, list4, listNames, filename, Table1, 
    Table2, Table3, Table4, ColName = "GeneSymbol", resDir = NULL) 
{
    require(Vennerable)
    require(colorfulVennPlot)
    require(RColorBrewer)
    cols <- c(brewer.pal(8, "Pastel1"), brewer.pal(8, "Pastel2"))
    cols <- cols[c(8, 2, 3, 15, 5, 6, 7, 1, 9, 10, 11, 12, 13, 
        14, 4, 16)]
    if (!is.null(resDir)) {
        resultsDir = resDir
    }
    list.venn <- list(list1, list2, list3, list4)
    names(list.venn) <- c(listNames[1], listNames[2], listNames[3], 
        listNames[4])
    vtest <- Venn(list.venn)
    vennData <- sapply(vtest@IntersectionSets, function(x) length(unlist(x)))
    pdf(file.path(resultsDir, paste("VennDiagram", filename, 
        "pdf", sep = ".")))
    plotVenn4d(vennData[-1], labels = c(listNames[1], listNames[2], 
        listNames[3], listNames[4]), Colors = cols, Title = "", 
        shrink = 0.8)
    dev.off()
    options(java.parameters = "-Xmx4g")
    require(XLConnect)
    xlcFreeMemory()
    wb <- loadWorkbook(file.path(resultsDir, paste("GeneLists", 
        filename, "xlsx", sep = ".")), create = TRUE)
    createSheet(wb, name = listNames[1])
    writeWorksheet(wb, Table1[Table1[, ColName] \%in\% unlist(vtest@IntersectionSets$`1111`), 
        ], sheet = listNames[1], startRow = 1, startCol = 1, 
        header = TRUE)
    createSheet(wb, name = listNames[2])
    writeWorksheet(wb, Table2[Table2[, ColName] \%in\% unlist(vtest@IntersectionSets$`1111`), 
        ], sheet = listNames[2], startRow = 1, startCol = 1, 
        header = TRUE)
    createSheet(wb, name = listNames[3])
    writeWorksheet(wb, Table3[Table3[, ColName] \%in\% unlist(vtest@IntersectionSets$`1111`), 
        ], sheet = listNames[3], startRow = 1, startCol = 1, 
        header = TRUE)
    createSheet(wb, name = listNames[4])
    writeWorksheet(wb, Table4[Table4[, ColName] \%in\% unlist(vtest@IntersectionSets$`1111`), 
        ], sheet = listNames[4], startRow = 1, startCol = 1, 
        header = TRUE)
    saveWorkbook(wb)
  }
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{ ~kwd1 }% use one of  RShowDoc("KEYWORDS")
\keyword{ ~kwd2 }% __ONLY ONE__ keyword per line
