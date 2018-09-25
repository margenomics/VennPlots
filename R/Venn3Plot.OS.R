Venn3Plot.OS <-
function(listG1, listG2, listG3, listNames, filename, data4T= NULL,
                      mkExcel = TRUE){
  #Funció per fer un venn diagram 3D
  #listG1: Vector of genes in group1
  #listG2: Vector of genes in group2
  #listG3: Vector of genes in group3
  #listNames: vector with the names of the lists of genes (normally contrasts)
  #filename: nom dels fitxers
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
  listG3 <- listG3[!is.na(listG3)]
  
  #Creem l'objecte del Venn
  list.venn<-list(listG1,listG2,listG3)
  names(list.venn)<-c(listNames[1],listNames[2],listNames[3])
  vtest<-Venn(list.venn)
  vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
  
  #PLOT VENN
  pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), 
             Colors=cols, Title="", shrink=1)
  dev.off()
  
  if(mkExcel){
    #EXCEL
    options( java.parameters = "-Xmx4g" ) #super important abans de cridar a XLConnect
    require(XLConnect)
    xlcFreeMemory()
    #cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
    cNames <- colnames(data4T)
      wb <- loadWorkbook(file.path(resultsDir,paste("GeneLists",filename,"xlsx",sep=".")),create=TRUE) #no li agraden els espais al nom o noms llargs!
      createSheet(wb, name = "orange") 
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`100`)),
                               cNames], 
                     sheet = "orange", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "blue")
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`010`)),
                               cNames], 
                     sheet = "blue", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green")
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`001`)),
                               cNames], 
                     sheet = "green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "pink")
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`110`)),
                               cNames], 
                     sheet = "pink", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "brown")
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`011`)),
                               cNames], 
                     sheet = "brown", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "yellow")
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`101`)),
                               cNames], 
                     sheet = "yellow", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Common grey")
      writeWorksheet(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`111`)),
                               cNames], 
                     sheet = "Common grey", startRow = 1, startCol = 1, header=TRUE)
      saveWorkbook(wb)
      
  }
}
