\name{VennPlot.3}
\alias{VennPlot.3}
\title{VennPlot.3}
\usage{
VennPlot.3(contrast.l = NULL , listNames, filename, data4T= NULL, symbols=TRUE,
   colnmes= c("AffyID", "Symbol"),listSymbols=F,listA=NULL, listB=NULL,listC=NULL,img.fmt= "pdf",
   pval=0, padj=0.05, logFC=1, up_down=FALSE)
}
\arguments{
\item{contrast.l}{List of contrasts which you want to use in order to perform a vennplot of them (ex: list(c("WT.vs.NKXX"),c("XXX.vs.WT_3"),c("RRR.vs.WT_5")))}
\item{listNames}{List of names for each section in vennplot (ex: c("WT.vs.NKXX","XXX.vs.WT_3","RRR.vs.WT_5"))}
\item{filename}{Name of output file (ex: "Results_Project_20200404")}
\item{data4T}{data4Tyers object. Default = NULL. If you want to filter each contrast using internal filtres you have to provide the data4Tyers object. If you don't have data4Tyers object (data4T=NULL) you must to use listSymbols=T}
\item{symbols}{symbols = T if genes name column is Symbols. Otherwise, symbols = F and you must indicated the column name for genes in "colnmes"}
\item{colnmes}{By default = c("AffyID","Symbol"). You must change this variable if you have another column name for genes}
\item{listSymbols}{By default = F. If you want to compare two lists of genes, pathways , ... -> listSymbols=T}
\item{listA}{Elements of list 1 (ex: c(pathway1,pathway2,pathway3))}
\item{listB}{Elements of list 2 (ex: c(pathway4, pathay5,pathway6))}
\item{listC}{Elements of list 3 (ex: c(pathway7, pathay8,pathway9))}
\item{img.fmt}{By default = pdf. The image format. It can be : "pdf" or "png"}
\item{pval}{By default = 0. If pval !=0 -> Uses pval value in order to filter the data}
\item{padj}{By default = 0.05. Uses padj value in order to filter the data}
\item{logFC}{By default = 1. Uses logfc value in order to filter the data}
\item{up_down}{By default = F. If it is TRUE, vennplots and excels will be done separately (up and down)}
}
\description{
Create Vennplot for 3 lists and an excel summary of the results
}
\examples{
VennPlot.3(contrast.l=list(c("WT.vs.NKXX"),c("XXX.vs.WT_3"),c("RRR.vs.WT_5")),listNames=c("WT.vs.NKXX","XXX.vs.WT_3","RRR.vs.WT_5"),filename="Results_Project_20200404",data4T=data4tyers, symbols=T, colnmes=c("AffyID","Symbol"), img.fmt="png",pval=0,padj=0.05,logFC=0.585, up_down=T)
}
