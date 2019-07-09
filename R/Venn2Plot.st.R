Venn2Plot.st <-
function(list1, list2, listNames, filename, Table1, Table2, 
                      ColName= "GeneSymbol", resDir=NULL){
  
  #list1: Vector amb els Gene symbols de una de les llistes
  #list2: Vector amb els Gene symbols de una de les llistes
  #listNames: Nom dels mÃ¨todes que corresponen a cada una de les llistes
  #filename: Nom del output file
  #Table1: Dataset del primer mÃ¨tode
  #Table2: Dataset del segon mÃ¨tode
  
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)  
  #establim els colors per als plots
  cols <- brewer.pal(8,"Pastel2") 
  
  if (!is.null(resDir)){
    resultsDir=resDir
  }
  
  #Ens assegurem que no hi ha NA
  list1 <- list1[!is.na(list1)]
  list2 <- list2[!is.na(list2)]
  
  #Creem l'objecte del Venn
  list.venn<-list(list1,list2)
  names(list.venn)<-c(listNames[1],listNames[2])
  vtest<-Venn(list.venn)
  vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
  
  #PLOT VENN
  #els noms estan al reves al plot labels
  if(img.fmt == "png") {
    png(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  } else if (img.fmt == "pdf"){
    pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  }
  plotVenn2d(vennData, labels=c(listNames[2],listNames[1]), Colors=cols, Title="", shrink=1)
  dev.off()
  
  #EXCEL
  options( java.parameters = "-Xmx4g" ) #super important abans de cridar a XLConnect
  require(XLConnect)
  xlcFreeMemory()
  
  wb <- loadWorkbook(file.path(resultsDir,paste("GeneLists",filename,"xlsx",sep=".")),create=TRUE) #no li agraden els espais al nom o noms llargs!
  createSheet(wb, name = listNames[1]) 
  writeWorksheet(wb,Table1[Table1[,ColName] %in% unlist(vtest@IntersectionSets$`11`),], 
                 sheet = listNames[1], startRow = 1, startCol = 1, header=TRUE)
  createSheet(wb, name = listNames[2])
  writeWorksheet(wb,Table2[Table2[,ColName] %in% unlist(vtest@IntersectionSets$`11`),], 
                 sheet = listNames[2], startRow = 1, startCol = 1, header=TRUE)
  saveWorkbook(wb)
  
}
