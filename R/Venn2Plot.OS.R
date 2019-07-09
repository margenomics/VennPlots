Venn2Plot.OS <-
function(listG1, listG2, listNames, filename, data4T= NULL, mkExcel = TRUE, img.fmt = "pdf"){
  #Funció per fer un venn diagram 2D
  #listG1: Vector amb els gene sym de la llista 1
  #listG2: Vector amb els gene sym de la llista 2
  #listNames: vector with the names of the lists of genes (normally contrasts)
  #filename: Nom dels fitxers
  #data4T: Data4Tyers
  #mkExcel: if TRUE an excel is made with the results of the venn diagram
  
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)  
  #establim els colors per als plots
  
  cols <- brewer.pal(8,"Pastel2") 
  
  #Ens assegurem de que no hi ha NA
  listG1 <- listG1[!is.na(listG1)]
  listG2 <- listG2[!is.na(listG2)]
  
  #Creem l'objecte del Venn
  list.venn<-list(listG1,listG2)
  names(list.venn)<-c(listNames[1],listNames[2])
  vtest<-Venn(list.venn)
  vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
  
  #PLOT VENN
  #els noms estan al reves al plot labels
  if(img.fmt == "png") {
    png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
  } else if (img.fmt == "pdf"){
    pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  }
  plotVenn2d(vennData, labels=c(listNames[2],listNames[1]), Colors=cols, Title="", shrink=1)
  dev.off()
  
  if(mkExcel) {
    #EXCEL
    options( java.parameters = "-Xmx4g" ) #super important abans de cridar a XLConnect
    require(XLConnect)
    xlcFreeMemory()
    #cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
    cNames <- colnames(data4T)
      wb <- loadWorkbook(file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),create=TRUE)
      createSheet(wb, name = listNames[1])
      writeWorksheet(wb,data4T[(!is.na(data4T[,"Symbol"])) & 
                                 (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`10`)), cNames], 
                     sheet = listNames[1], startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = listNames[2])
      writeWorksheet(wb,data4T[(!is.na(data4T[,"Symbol"])) & 
                                 (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`01`)), cNames], 
                     sheet = listNames[2], startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Common")
      writeWorksheet(wb,data4T[(!is.na(data4T[,"Symbol"])) & 
                                 (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`11`)),cNames],
                     sheet = "Common", startRow = 1, startCol = 1, header=TRUE)
      saveWorkbook(wb)
      
  }
}
