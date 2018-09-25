Venn3Plot.st <-
function(list1, list2, list3, listNames, filename, Table1, Table2, 
                      Table3, ColName= "GeneSymbol", resDir=NULL ){
 
  #list1: Vector amb els Gene symbols de una de les llistes
  #list2: Vector amb els Gene symbols de una de les llistes
  #list3: Vector amb els Gene symbols de una de les llistes
  #listNames: Nom dels mÃ¨todes que corresponen a cada una de les llistes
  #filename: Nom del output file
  #Table1: Dataset del primer mÃ¨tode
  #Table2: Dataset del segon mÃ¨tode
  #Table3: Dataset del tercer mÃ¨tode
  
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)  
  #establim els colors per als plots
  cols <- brewer.pal(8,"Pastel2") 
  
  if (!is.null(resDir)){
    resultsDir=resDir
  }
  
  #Mirem que no hi hagi NA
  list1 <- list1[!is.na(list1)]
  list2 <- list2[!is.na(list2)]
  list3 <- list3[!is.na(list3)]
  
  #Creem l'objecte del Venn
  list.venn<-list(list1,list2,list3)
  names(list.venn)<-c(listNames[1],listNames[2],listNames[3])
  vtest<-Venn(list.venn)
  vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
  
  #PLOT VENN
  pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), 
             Colors=cols, Title="", shrink=1)
  dev.off()
  
  
  #EXCEL
  options( java.parameters = "-Xmx4g" ) #super important abans de cridar a XLConnect
  require(XLConnect)
  xlcFreeMemory()
    
  wb <- loadWorkbook(file.path(resultsDir,paste("GeneLists",filename,"xlsx",sep=".")),create=TRUE) #no li agraden els espais al nom o noms llargs!
  createSheet(wb, name = listNames[1]) 
  writeWorksheet(wb,Table1[Table1[,ColName] %in% unlist(vtest@IntersectionSets$`111`),], 
                     sheet = listNames[1], startRow = 1, startCol = 1, header=TRUE)
  createSheet(wb, name = listNames[2])
  writeWorksheet(wb,Table2[Table2[,ColName] %in% unlist(vtest@IntersectionSets$`111`),], 
                     sheet = listNames[2], startRow = 1, startCol = 1, header=TRUE)
  createSheet(wb, name = listNames[3])
  writeWorksheet(wb,Table3[Table3[,ColName] %in% unlist(vtest@IntersectionSets$`111`),], 
                     sheet = listNames[3], startRow = 1, startCol = 1, header=TRUE)
  saveWorkbook(wb)
   
}
