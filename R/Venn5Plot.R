Venn5Plot <-
function(listG1,listG2,listG3,listG4,listG5,listNames,filename,data4T= NULL,symbols=TRUE,
                      mkExcel = TRUE,colnmes= c("AffyID", "Symbol"), CatCex=0.8, CatDist=rep(0.1, 5)){ 
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)

  cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
  cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]      ##Reorganitzar els colors per a que no coincideixin tonalitats semblants
  
  #Creem l'objecte del Venn
  if (symbols){
    list.1<-listG1[,colnmes[2]]
    list.2<-listG2[,colnmes[2]]
    list.3<-listG3[,colnmes[2]]
    list.4<-listG4[,colnmes[2]]
    list.5<-listG5[,colnmes[2]]
  }else{
    list.1<-listG1[,colnmes[1]]
    list.2<-listG2[,colnmes[1]]
    list.3<-listG3[,colnmes[1]]
    list.4<-listG4[,colnmes[1]]
    list.5<-listG5[,colnmes[1]]
  }
  
  list.venn<-list(list.1,list.2,list.3,list.4,list.5)

  listNames=c(name.1,name.2,name.3,name.4,name.5)
  names(list.venn)<-listNames
  vtest<-Venn(list.venn)
  
  require(VennDiagram)
  pdf(file.path(resultsDir,paste("VennDiagram",filename, "pdf",sep=".")))
  venn.plot <- draw.quintuple.venn(
    area1 = length(list.1),
    area2 = length(list.2),
    area3 = length(list.3),
    area4 = length(list.4),
    area5 = length(list.5),
    n12 = length(intersect(list.1,list.2)),
    n13 = length(intersect(list.1,list.3)),
    n14 = length(intersect(list.1,list.4)),
    n15 = length(intersect(list.1,list.5)),
    n23 = length(intersect(list.2,list.3)),
    n24 = length(intersect(list.2,list.4)),
    n25 = length(intersect(list.2,list.5)),
    n34 = length(intersect(list.3,list.4)),
    n35 = length(intersect(list.3,list.5)),
    n45 = length(intersect(list.4,list.5)),
    n123 = length(intersect(list.1,intersect(list.2,list.3))),
    n124 = length(intersect(list.1,intersect(list.2,list.4))),
    n125 = length(intersect(list.1,intersect(list.2,list.5))),
    n134 = length(intersect(list.1,intersect(list.3,list.4))),
    n135 = length(intersect(list.1,intersect(list.3,list.5))),
    n145 = length(intersect(list.1,intersect(list.4,list.5))),
    n234 = length(intersect(list.2,intersect(list.3,list.4))),
    n235 = length(intersect(list.2,intersect(list.3,list.5))),
    n245 = length(intersect(list.2,intersect(list.4,list.5))),
    n345 = length(intersect(list.3,intersect(list.4,list.5))),
    n1234 = length(intersect(list.1,intersect(list.2,intersect(list.3,list.4)))),
    n1235 = length(intersect(list.1,intersect(list.2,intersect(list.3,list.5)))),
    n1245 = length(intersect(list.1,intersect(list.2,intersect(list.4,list.5)))),
    n1345 = length(intersect(list.1,intersect(list.3,intersect(list.4,list.5)))),
    n2345 = length(intersect(list.2,intersect(list.3,intersect(list.4,list.5)))),
    n12345 = length(intersect(list.1,intersect(list.2,intersect(list.3,intersect(list.4,list.5))))),
    category = listNames,
    fill = c("dodgerblue", "goldenrod1", "darkorange1", "seagreen3", "orchid3"),
    cat.cex = CatCex,
    cat.dist = CatDist,
    margin = 0.05,
    cex = c(1.5, 1.5, 1.5, 1.5, 1.5, 1, 0.8, 1, 0.8, 1, 0.8, 1, 0.8, 1, 0.8,
            1, 0.55, 1, 0.55, 1, 0.55, 1, 0.55, 1, 0.55, 1, 1, 1, 1, 1, 1.5),
    ind = TRUE)
  
  dev.off()
  
  if(mkExcel) {
    #EXCEL
    options( java.parameters = "-Xmx4g" ) #super important abans de cridar a XLConnect
    require(XLConnect)
    xlcFreeMemory()
    cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
    wb <- loadWorkbook(file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),create=TRUE)
    if (symbols) {
      createSheet(wb, name = "Blue")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10000`)) & 
                                 (data4T[,colnmes[1]] %in% listG1[,colnmes[1]]), cNames], 
                     sheet = "Blue", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01000`)) & 
                                 (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]), cNames], 
                     sheet = "Yellow", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Orange")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00100`)) & 
                                 (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]),cNames],
                     sheet = "Orange", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00010`)) & 
                                 (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]),cNames],
                     sheet = "Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00001`)) & 
                                 (data4T[,colnmes[1]] %in% listG5[,colnmes[1]]),cNames],
                     sheet = "Purple", startRow = 1, startCol = 1, header=TRUE)
      ###############################################################################################
      createSheet(wb, name = "Blue.Yellow")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11000`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Orange")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10100`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]])),cNames],
                     sheet = "Blue.Orange", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10010`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Blue.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10001`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Orange")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01100`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]])),cNames],
                     sheet = "Yellow.Orange", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01010`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Yellow.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01001`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Yellow.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Orange.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00110`)) & 
                                 ((data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Orange.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Orange.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00101`)) & 
                                 ((data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00011`)) & 
                                 ((data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      ##############################################################################################
      createSheet(wb, name = "Blue.Yellow.Orange")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11100`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow.Orange", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Yellow.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11010`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Yellow.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11001`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Orange.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10110`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Blue.Orange.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Orange.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10101`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10011`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Orange.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01110`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Yellow.Orange.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Orange.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01101`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Yellow.Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01011`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Yellow.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Orange.Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`00111`)) & 
                                 ((data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Orange.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      ############################################################################################
      createSheet(wb, name = "Blue.Yellow.Orange.Green")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11110`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow.Orange.Green", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Orange.Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10111`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Orange.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Yellow.Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11011`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Blue.Yellow.Orange.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11101`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Blue.Yellow.Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Yellow.Orange.Green.Purple")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01111`)) & 
                                 ((data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Yellow.Orange.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
      createSheet(wb, name = "Common.All")
      writeWorksheet(wb,data4T[(!is.na(data4T[,colnmes[2]])) & 
                                 (data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11111`)) & 
                                 ((data4T[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                    (data4T[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG4[,colnmes[1]]) |
                                    (data4T[,colnmes[1]] %in% listG5[,colnmes[1]])),cNames],
                     sheet = "Common.All", startRow = 1, startCol = 1, header=TRUE)
      
      } else {
      
        createSheet(wb, name = "Blue")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10000`),cNames], 
                       sheet = "Blue", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01000`),cNames], 
                       sheet = "Yellow", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Orange")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00100`),cNames],
                       sheet = "Orange", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00010`),cNames],
                       sheet = "Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00001`),cNames],
                       sheet = "Purple", startRow = 1, startCol = 1, header=TRUE)
        ###############################################################################################
        createSheet(wb, name = "Blue.Yellow")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11000`),cNames],
                       sheet = "Blue.Yellow", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Orange")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10100`),cNames],
                       sheet = "Blue.Orange", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10010`),cNames],
                       sheet = "Blue.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10001`),cNames],
                       sheet = "Blue.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Orange")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01100`),cNames],
                       sheet = "Yellow.Orange", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01010`),cNames],
                       sheet = "Yellow.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01001`),cNames],
                       sheet = "Yellow.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Orange.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00110`),cNames],
                       sheet = "Orange.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Orange.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00101`),cNames],
                       sheet = "Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00011`),cNames],
                       sheet = "Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        ##############################################################################################
        createSheet(wb, name = "Blue.Yellow.Orange")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11100`),cNames],
                       sheet = "Blue.Yellow.Orange", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Yellow.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11010`),cNames],
                       sheet = "Blue.Yellow.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Yellow.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11001`),cNames],
                       sheet = "Blue.Yellow.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Orange.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10110`),cNames],
                       sheet = "Blue.Orange.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Orange.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10101`),cNames],
                       sheet = "Blue.Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10011`),cNames],
                       sheet = "Blue.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Orange.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01110`),cNames],
                       sheet = "Yellow.Orange.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Orange.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01101`),cNames],
                       sheet = "Yellow.Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01011`),cNames],
                       sheet = "Yellow.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Orange.Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`00111`),cNames],
                       sheet = "Orange.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        ############################################################################################
        createSheet(wb, name = "Blue.Yellow.Orange.Green")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11110`),cNames],
                       sheet = "Blue.Yellow.Orange.Green", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Orange.Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10111`),cNames],
                       sheet = "Blue.Orange.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Yellow.Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11011`),cNames],
                       sheet = "Blue.Yellow.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Blue.Yellow.Orange.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11101`),cNames],
                       sheet = "Blue.Yellow.Orange.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Yellow.Orange.Green.Purple")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01111`),cNames],
                       sheet = "Yellow.Orange.Green.Purple", startRow = 1, startCol = 1, header=TRUE)
        createSheet(wb, name = "Common.All")
        writeWorksheet(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11111`),cNames],
                       sheet = "Common.All", startRow = 1, startCol = 1, header=TRUE)
    }
    saveWorkbook(wb)
  }
  
}
