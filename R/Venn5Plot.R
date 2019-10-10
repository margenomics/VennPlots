Venn5Plot <- function(list.1,list.2,list.3,list.4,list.5,listNames,filename,data4T= NULL,
           img.fmt = "pdf", mkExcel = TRUE,colnmes= "Symbol", CatCex=0.8, CatDist=rep(0.1, 5)){ 
  
    #11/10/19 use 'openxlsx' package to write xlsx files, no java dependencies
  
    require(Vennerable) 
    require(colorfulVennPlot) #Per generar plots amb colors diferents
    require(RColorBrewer)
    require(openxlsx)
    
    cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
    cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]      ##Reorganitzar els colors per a que no coincideixin tonalitats semblants
    
    #Creem l'objecte del Venn
    list.venn<-list(list.1,list.2,list.3,list.4,list.5)
    
    names(list.venn)<-listNames
    vtest<-Venn(list.venn)
    
    require(VennDiagram)
    if(img.fmt == "png") {
           png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
    } else if (img.fmt == "pdf"){
           pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
    }
    
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
      ind = F)
    
    dev.off()
    
    if(mkExcel) {
      #EXCEL
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      wb <- createWorkbook()
   
      addWorksheet(wb, sheetName = "Blue")
      writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10000`),], 
                       sheet = "Blue", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Yellow")
      writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01000`),], 
                       sheet = "Yellow", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Orange")
      writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00100`),],
                       sheet = "Orange", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Green")
      writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00010`),],
                       sheet = "Green", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Purple")
      writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00001`),],
                       sheet = "Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        ###############################################################################################
        addWorksheet(wb, sheetName = "Blue.Yellow")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11000`),],
                       sheet = "Blue.Yellow", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Orange")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10100`),],
                       sheet = "Blue.Orange", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10010`),],
                       sheet = "Blue.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10001`),],
                       sheet = "Blue.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Orange")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01100`),],
                       sheet = "Yellow.Orange", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01010`),],
                       sheet = "Yellow.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01001`),],
                       sheet = "Yellow.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Orange.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00110`),],
                       sheet = "Orange.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Orange.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00101`),],
                       sheet = "Orange.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00011`),],
                       sheet = "Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        ##############################################################################################
        addWorksheet(wb, sheetName = "Blue.Yellow.Orange")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11100`),],
                       sheet = "Blue.Yellow.Orange", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Yellow.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11010`),],
                       sheet = "Blue.Yellow.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Yellow.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11001`),],
                       sheet = "Blue.Yellow.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Orange.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10110`),],
                       sheet = "Blue.Orange.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Orange.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10101`),],
                       sheet = "Blue.Orange.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10011`),],
                       sheet = "Blue.Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Orange.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01110`),],
                       sheet = "Yellow.Orange.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Orange.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01101`),],
                       sheet = "Yellow.Orange.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01011`),],
                       sheet = "Yellow.Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Orange.Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`00111`),],
                       sheet = "Orange.Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        ############################################################################################
        addWorksheet(wb, sheetName = "Blue.Yellow.Orange.Green")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11110`),],
                       sheet = "Blue.Yellow.Orange.Green", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Orange.Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`10111`),],
                       sheet = "Blue.Orange.Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Yellow.Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11011`),],
                       sheet = "Blue.Yellow.Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Blue.Yellow.Orange.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11101`),],
                       sheet = "Blue.Yellow.Orange.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Yellow.Orange.Green.Purple")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`01111`),],
                       sheet = "Yellow.Orange.Green.Purple", startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common.All")
        writeData(wb,data4T[data4T[,colnmes] %in% unlist(vtest@IntersectionSets$`11111`),],
                       sheet = "Common.All", startRow = 1, startCol = 1, headerStyle = hs1)
      
      saveWorkbook(wb,file=file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
    }
    
  }
