# Venn Diagrams for genomic data

## Introduction

The aim of this R package is to easy plot the intersections of different gene lists in venn diagrams. It also creates tables for genes in each region together with their annotation. The functions that can be found in the package are:


* **VennPlot.2** : Function to create Venn diagrams from two tables or lists containing the columns with array ids or gene symbols, ideally both of them. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained.  
* **VennPlot.3**: Function to create Venn diagrams from three tables or lists containing the columns with array ids or gene symbols, ideally both of them. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained. 
* **VennPlot.4** : Function to create Venn diagrams from four tables or lists containing the columns with array ids or gene symbols, ideally both of them. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained.
* **Venn5Plot** : Function to make Venn diagrams from five tables containing the columns with array ids or gene symbols, ideally both of them. Results are shown in a plot, and the lists of genes corresponding to each region can be obtained.


## Installation

         library(devtools)
         devtools::install_github("margenomics/VennPlots")
         library(VennPlots)

## Usage

        # VennPlots for 2 lists or contrast  
        VennPlot.2(contrast.l=list(c("WT.vs.NKXX"),c("XXX.vs.WT_3")),listNames=c("WT.vs.NKXX","XXX.vs.WT_3"),
                   filename="Results_Project_20200404",data4T=data4tyers, symbols=T, colnmes=c("AffyID","Symbol"), 
                   img.fmt="png",pval=0,padj=0.05,logFC=0.585, up_down=T)  
                     
        # VennPlots for 3 lists or contrast  
        VennPlot.3(contrast.l=list(c("WT.vs.NKXX"),c("XXX.vs.WT_3"),c("RRR.vs.WT_5")),
                   listNames=c("WT.vs.NKXX","XXX.vs.WT_3","RRR.vs.WT_5"),
                   filename="Results_Project_20200404",data4T=data4tyers, 
                   symbols=T, colnmes=c("AffyID","Symbol"), img.fmt="png",
                   pval=0,padj=0.05,logFC=0.585, up_down=T)  
                     
        # VennPlots for 4 lists or contrast  
        VennPlot.4(contrast.l=list(c("WT.vs.NKXX"),c("XXX.vs.WT_3"),c("RRR.vs.WT_5"),c("WT_2.vs.FRR")),
                   listNames=c("WT.vs.NKXX","XXX.vs.WT_3","RRR.vs.WT_5","WT_2.vs.FRR"),
                   filename="Results_Project_20200404",data4T=data4tyers, symbols=T, 
                   colnmes=c("AffyID","Symbol"), img.fmt="png",pval=0,padj=0.05,logFC=0.585, up_down=T)  
                     
        # VennPlots for 5 lists or contrast  
        Venn5Plot(list.1,list.2,list.3,list.4,list.5,listNames,filename,data4T= NULL,
                  img.fmt = "pdf", mkExcel = TRUE,colnmes= "Symbol", CatCex=0.8, CatDist=rep(0.1, 5))  
  
  
## Example 

#### VennPlot.2

          library(VennPlots)
          resultsDir = "/home/user/project/results"
          listA = c("ADCY5","APOC3","COQ7","EMD","ERBB2","HBP1","IL2", "IL6", "IL8", "MTOR", "NBN","PCNA", "PEX5","SP1", "STAT3")
          listB = c("ABL1", "APOC3", "BCL2", "BRCA1", "EMD", "IL2", "IL6", "LMNA","MAPK3", "NGF", "PCK1", "PLAU", "SIRT1", "XPA", "XRCC5" )
            
          VennPlot.2(contrast.l = NULL,listA= listA, listB= listB, listNames=c("WT.vs.NKXXa","WT.vs.NKXXb"),
          filename="Results_Project_WT.vs.NKXXab",data4T=NULL, listSymbols=T, img.fmt="png")

##### **_Excel_**:
**WT.vs.NKXXa**  
ADCY5  
COQ7  
ERBB2  
HBP1  
IL8  
MTOR  
NBN  
PCNA  
PEX5  
SP1  
STAT3  
  
**WT.vs.NKXXb**  
ABL1  
BCL2  
BRCA1  
LMNA  
MAPK3  
NGF  
PCK1  
PLAU  
SIRT1  
XPA  
XRCC5  
  
**Common**  
APOC3  
EMD  
IL2  
IL6  

##### **_Plot_**:

![Venn2Plot](https://github.com/margenomics/VennPlots/blob/master/images/VennDiagram.Results_Project_WT.vs.NKXXab.png)


#### VennPlot.3

          library(VennPlots)
          resultsDir = "/home/user/project/results"
          listA = c("ADCY5","APOC3","COQ7","EMD","ERBB2","HBP1","IL2", "IL6", "IL8", "MTOR", "NBN","PCNA", "PEX5","SP1", "STAT3")
          listB = c("ABL1", "APOC3", "BCL2", "BRCA1", "EMD", "IL2", "IL6", "LMNA","MAPK3", "NGF", "PCK1", "PLAU", "SIRT1", "XPA", "XRCC5" )
          listC =  c("ABL1", "APEX1","ATF2","IL2","IL6", "MAPK3","NBN", "PPARG","PROP1","PYCR1","SHC1","SP1", "TRAP1", "UCP2", "XPA")  
          
          VennPlot.3(contrast.l = NULL,listA= listA, listB= listB, listC=listC, listNames=c("WT.vs.NKXXa","WT.vs.NKXXb","WT.vs.NKXXc"),
          filename="Results_Project_WT.vs.NKXXabc",data4T=NULL, listSymbols=T, img.fmt="png")

##### **_Excel_**:
**Orange**  
ADCY5  
COQ7  
ERBB2  
HBP1  
IL8  
MTOR  
PCNA  
PEX5  
STAT3  
  
**Blue**  
BCL2  
BRCA1  
LMNA  
NGF  
PCK1  
PLAU  
SIRT1  
XRCC5  
  
**Green**  
APEX1  
ATF2  
PPARG  
PROP1  
PYCR1  
SHC1  
TRAP1  
UCP2  
  
**Pink**  
APOC3  
EMD  
  
**Brown**  
ABL1   
MAPK3  
XPA  
  
**Yellow**  
NBN  
SP1  
  
**Common(Grey)**  
IL2  
IL6  
  
##### **_Plot_**:

![Venn3Plot](https://github.com/margenomics/VennPlots/blob/master/images/VennDiagram.Results_Project_WT.vs.NKXXabc.png)





#### VennPlot.4

          library(VennPlots)
          resultsDir = "/home/user/project/results"
          listA = c("ADCY5","APOC3","COQ7","EMD","ERBB2","HBP1","IL2", "IL6", "IL8", "MTOR", "NBN","PCNA", "PEX5","SP1", "STAT3")
          listB = c("ABL1", "APOC3", "BCL2", "BRCA1", "EMD", "IL2", "IL6", "LMNA","MAPK3", "NGF", "PCK1", "PLAU", "SIRT1", "XPA", "XRCC5" )
          listC =  c("ABL1", "APEX1","ATF2","IL2","IL6", "MAPK3","NBN", "PPARG","PROP1","PYCR1","SHC1","SP1", "TRAP1", "UCP2", "XPA")  
          listD = c("ABL1","ADCY5", "APOC3", "COQ7","EMD","MAPK3","IL2","IL6", "NGF","PROP1", "PLAU","TRAP1","UCP3","VEGFA", "YWHAZ")
          
          VennPlot.4(contrast.l = NULL,listA= listA, listB= listB, listC=listC, 
          listD=listD,listNames=c("WT.vs.NKXXa","WT.vs.NKXXb","WT.vs.NKXXc","WT.vs.NKXXd"),
          filename="Results_Project_WT.vs.NKXXabcd",data4T=NULL, listSymbols=T, img.fmt="png")

##### **_Excel_**:  
  
**Pink**  
ERBB2  
HBP1  
IL8  
MTOR  
PCNA  
PEX5  
STAT3  
  
**Blue2**  
BCL2  
BRCA1  
LMNA  
PCK1  
SIRT1  
XRCC5  
  
**Green2**  
  
**Brown1**  
APEX1  
ATF2  
PPARG  
PYCR1  
SHC1  
UCP2  
  
**Orange1**  
NBN  
SP1  
  
**Yellow1**  
XPA  
  
**Brown2**  
  
**Red**  
UCP3  
VEGFA  
YWHAZ  
  
**Green3**   
ADCY5  
COQ7   
  
**Orange2**  
NGF  
PLAU  
  
**Blue1**  
APOC3  
EMD  
  
**Purple1**  
PROP1  
TRAP1  

**Green1** 
  
**Yellow2**   
ABL1     
MAPK3  
  
**Purple2**  
IL2  
IL6  
  
##### **_Plot_**:

![Venn4Plot](https://github.com/margenomics/VennPlots/blob/master/images/VennDiagram.Results_Project_WT.vs.NKXXabcd.png)


