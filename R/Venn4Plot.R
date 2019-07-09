Venn4Plot <-
function(listG1, listG2, listG3, listG4, listNames, filename, data4T= NULL, img.fmt = "pdf",
                      symbols=TRUE, mkExcel = TRUE, colnmes= c("AffyID", "Symbol")){
  #Funció per fer un venn diagram 3D
  #listG1: Data frame amb els resultats del data4Tyers del primer grup que volem comparar
  #listG2: Data frame amb els resultats del data4Tyers del segon grup que volem comparar
  #listG3: Data frame amb els resultats del data4Tyers del tercer grup que volem comparar
  #listG4: Data frame amb els resultats del data4Tyers del quart grup que volem comparar
  #listNames: vector with the names of the lists of genes (normally contrasts)
  #filename: Nom dels grafics
  #data4T: quan mkExcel = TRUE l'objecte Data4Tyers amb el que es fara el excel
  #symbols: if TRUE gene symbols are used to make the venn 
  #mkExcel: if TRUE an excel is made with the results of the venn diagram
  #colnmes: Nom de les columnes AffyID i Symbol
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)
  #establim els colors per als plots
  cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
  cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]      ##Reorganitzar els colors per a que no coincideixin tonalitats semblants
  
  #Creem l'objecte del Venn
  if (symbols){
    listG1<-listG1[!is.na(listG1[,colnmes[2]]),]
    listG2<-listG2[!is.na(listG2[,colnmes[2]]),]
    listG3<-listG3[!is.na(listG3[,colnmes[2]]),]
    listG4<-listG4[!is.na(listG4[,colnmes[2]]),]
    list1<-listG1[,colnmes[2]]
    list2<-listG2[,colnmes[2]]
    list3<-listG3[,colnmes[2]]
    list4<-listG4[,colnmes[2]]
  }else{
    list1<-listG1[,colnmes[1]]
    list2<-listG2[,colnmes[1]]
    list3<-listG3[,colnmes[1]]
    list4<-listG4[,colnmes[1]]
  }
  
  list.venn<-list(list1,list2,list3, list4)
  names(list.venn)<-c(listNames[1],listNames[2],listNames[3], listNames[4])
  vtest<-Venn(list.venn)
  vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
  
  #PLOT VENN
  if(img.fmt == "png") {
    png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
  } else if (img.fmt == "pdf"){
    pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
  }
  plotVenn4d(vennData[-1], labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
             Colors=cols, Title="", shrink=0.8)
  dev.off()
  
  if(mkExcel){
    #EXCEL
    options( java.parameters = "-Xmx4g" ) #super important abans de cridar a XLConnect
    require(XLConnect)
    xlcFreeMemory()
    cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
    
    if(symbols){
      wb <- loadWorkbook(paste(resultsDir,
                                paste("GeneLists",filename,"xlsx",sep="."),
                                sep="/"),create=TRUE)
      
      createSheet(wb, name = "pink")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2])) & 
                                 (data4T[,colnmes[1]] %in% listG1[,colnmes[1]]), cNames], 
                     sheet = "pink", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "blue2")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3])) & 
                                 (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]), cNames], 
                     sheet = "blue2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green2")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4])) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) | 
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])), cNames], 
                     sheet = "green2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "brown1")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5])) & 
                                 (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]), cNames], 
                     sheet = "brown1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "orange1")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6])) & 
                                 ((data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG1[,colnmes[1]])), cNames], 
                     sheet = "orange1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "yellow1")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7])) & 
                                 ((data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG2[,colnmes[1]])), cNames], 
                     sheet = "yellow1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "brown2")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8])) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG2[,colnmes[1]])), cNames], 
                     sheet = "brown2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "red")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9])) & 
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]), cNames], 
                     sheet = "red", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green3")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10])) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) | 
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "green3", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "orange2")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11])) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "orange2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "blue1")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12])) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) | 
                                 (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "blue1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "purple1")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13])) & 
                                 ((data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "purple1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green1")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14])) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "green1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "yellow2")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15])) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "yellow2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "purple2")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16])) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                     sheet = "purple2", startRow = 1, startCol = 1, header=TRUE)
      
      saveWorkbook(wb)
    }else{
      wb <- loadWorkbook(paste(resultsDir,
                                paste("GeneLists",filename,"xlsx",sep="."),
                                sep="/"),create=TRUE)
      
      createSheet(wb, name = "pink")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[2]),cNames], 
                     sheet = "pink", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "blue2")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[3]),cNames], 
                     sheet = "blue2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green2")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[4]),cNames], 
                     sheet = "green2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "brown1")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[5]),cNames], 
                     sheet = "brown1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "orange1")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[6]),cNames], 
                     sheet = "orange1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "yellow1")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[7]),cNames], 
                     sheet = "yellow1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "brown2")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[8]),cNames], 
                     sheet = "brown2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "red")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[9]),cNames], 
                     sheet = "red", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green3")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[10]),cNames], 
                     sheet = "green3", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "orange2")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[11]),cNames], 
                     sheet = "orange2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "blue1")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[12]),cNames], 
                     sheet = "blue1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "purple1")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[13]),cNames], 
                     sheet = "purple1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "green1")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[14]),cNames], 
                     sheet = "green1", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "yellow2")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[15]),cNames], 
                     sheet = "yellow2", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "purple2")
      writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[16]),cNames], 
                     sheet = "purple2", startRow = 1, startCol = 1, header=TRUE)
      
      saveWorkbook(wb)
    }
  }
}
