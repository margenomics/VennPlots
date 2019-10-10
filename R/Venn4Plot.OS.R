Venn4Plot.OS <-
function(listG1, listG2, listG3, listG4, listNames, filename, data4T= NULL, mkExcel = TRUE, img.fmt = "pdf"){
    #Funció per fer un venn diagram 3D
    #listG1: vector amb els gens o TCI DE primer grup que volem comparar
    #listG2: vector amb els gens o TCI DE del segon grup que volem comparar
    #listG3: vector amb els gens o TCI DE del tercer grup que volem comparar
    #listG4: vector amb els gens o TCI DE del quart grup que volem comparar
    #listNames: vector with the names of the lists of genes (normally contrasts)
    #filename: Nom dels fitxers
    #data4T: quan mkExcel = TRUE l'objecte Data4Tyers amb el que es fara el excel
    #mkExcel: if TRUE an excel is made with the results of the venn diagram
   #11/10/19 use 'openxlsx' package to write xlsx files, no java dependencies
  
    require(Vennerable) 
    require(colorfulVennPlot) #Per generar plots amb colors diferents
    require(RColorBrewer)
    require(openxlsx)
    #establim els colors per als plots
    cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
    cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]      ##Reorganitzar els colors per a que no coincideixin tonalitats semblants
    
    #Ens assegurem de que no hi ha NA
    listG1 <- listG1[!is.na(listG1)]
    listG2 <- listG2[!is.na(listG2)]
    listG3 <- listG3[!is.na(listG3)]
    listG4 <- listG4[!is.na(listG4)]
    
    #Creem l'objecte del Venn
    list.venn<-list(listG1,listG2,listG3, listG4)
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
        cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
        hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
            wb <- createWorkbook()
            
            addWorksheet(wb, sheetName = "pink")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[2]),cNames], 
                           sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "blue2")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[3]),cNames], 
                           sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "green2")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[4]),cNames], 
                           sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "brown1")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[5]),cNames], 
                           sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "orange1")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[6]),cNames], 
                           sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "yellow1")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[7]),cNames], 
                           sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "brown2")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[8]),cNames], 
                           sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "red")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[9]),cNames], 
                           sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "green3")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[10]),cNames], 
                           sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "orange2")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[11]),cNames], 
                           sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "blue1")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[12]),cNames], 
                           sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "purple1")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[13]),cNames], 
                           sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "green1")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[14]),cNames], 
                           sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "yellow2")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[15]),cNames], 
                           sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
            addWorksheet(wb, sheetName = "purple2")
            writeData(wb,data4T[data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets[16]),cNames], 
                           sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
            
            saveWorkbook(wb,file=paste(resultsDir,
                                       paste("GeneLists",filename,"xlsx",sep="."),
                                       sep="/"),overwrite = TRUE)
    }
}
