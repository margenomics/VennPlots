Venn3Plot.OS <-
function(listG1, listG2, listG3, listNames, filename, data4T= NULL,
                      mkExcel = TRUE, img.fmt = "pdf"){
  #Funció per fer un venn diagram 3D
  #listG1: Vector of genes in group1
  #listG2: Vector of genes in group2
  #listG3: Vector of genes in group3
  #listNames: vector with the names of the lists of genes (normally contrasts)
  #filename: nom dels fitxers
  #data4T: Data4Tyers
  #mkExcel: if TRUE an excel is made with the results of the venn diagram
  #11/10/19 use 'openxlsx' package to write xlsx files, no java dependencies
  
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)  
  require(openxlsx)
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
  if(img.fmt == "png") {
    png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
  } else if (img.fmt == "pdf"){
    pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  }
  plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), 
             Colors=cols, Title="", shrink=1)
  dev.off()
  
  if(mkExcel){
    #EXCEL
    #cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
    cNames <- colnames(data4T)
    hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                       border = "Bottom", fontColour = "white")
    
      wb <- createWorkbook()
      addWorksheet(wb, sheetName = "orange") 
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`100`)),
                               cNames], 
                     sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "blue")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`010`)),
                               cNames], 
                     sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "green")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`001`)),
                               cNames], 
                     sheet = "green", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "pink")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`110`)),
                               cNames], 
                     sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "brown")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`011`)),
                               cNames], 
                     sheet = "brown", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "yellow")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`101`)),
                               cNames], 
                     sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Common grey")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`111`)),
                               cNames], 
                     sheet = "Common grey", startRow = 1, startCol = 1, headerStyle = hs1)
      saveWorkbook(wb,file=file.path(resultsDir,paste("GeneLists",filename,"xlsx",sep=".")),overwrite = TRUE)
      
  }
}
