VennPlot.4 <-function(contrast.l = NULL , listNames, filename, data4T= NULL, img.fmt = "pdf",
                            symbols=TRUE, colnmes= c("AffyID", "Symbol"),listSymbols=F,
                            listA=NULL, listB=NULL,listC=NULL,listD=NULL,
                            pval=0, padj=0.05, logFC=1, up_down=FALSE){
  
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)
  require(openxlsx)
  
  if (listSymbols){
    #establim els colors per als plots
    cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
    cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]      ##Reorganitzar els colors per a que no coincideixin tonalitats semblants
    
    listG1 <- listA
    listG2 <- listB
    listG3 <- listC
    listG4 <- listD
    
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
      png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")),width = 1200, height = 700)
    } else if (img.fmt == "pdf"){
      pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 700)
    }
    plotVenn4d(vennData[-1], labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
               Colors=cols, Title="", shrink=0.5)
    dev.off()
    
    if(!is.null(data4T)){
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
                                 paste("VennGenes",filename,"xlsx",sep="."),
                                 sep="/"),overwrite = TRUE)
    }else{
      genes <- data.frame("Genes"=unlist(list.venn))
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      wb <- createWorkbook()
      
      addWorksheet(wb, sheetName = "pink")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[2]),]), 
                sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "blue2")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[3]),]), 
                sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "green2")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[4]),]), 
                sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "brown1")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[5]),]), 
                sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "orange1")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[6]),]), 
                sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "yellow1")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[7]),]), 
                sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "brown2")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[8]),]), 
                sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "red")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[9]),]), 
                sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "green3")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[10]),]), 
                sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "orange2")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[11]),]), 
                sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "blue1")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[12]),]), 
                sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "purple1")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[13]),]), 
                sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "green1")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[14]),]), 
                sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "yellow2")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[15]),]), 
                sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "purple2")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets[16]),]), 
                sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
      
      saveWorkbook(wb,file=paste(resultsDir,
                                 paste("VennGenes",filename,"xlsx",sep="."),
                                 sep="/"),overwrite = TRUE)
    }
    
  }else{
    
    if (up_down){
      if (pval!=0){
        listG1_up <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2_up <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        listG3_up <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[3],sep=".")] < pval,]
        listG4_up <- data4T[data4T[paste("logFC",contrast.l[4],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[4],sep=".")] < pval,]
        
        
        listG1_dn <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2_dn <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        listG3_dn <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[3],sep=".")] < pval,]
        listG4_dn <- data4T[data4T[paste("logFC",contrast.l[4],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[4],sep=".")] < pval,]
        
      }else{
        listG1_up <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2_up <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        listG3_up <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
        listG4_up <- data4T[data4T[paste("logFC",contrast.l[4],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[4],sep=".")] < padj,]
        
        
        listG1_dn <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2_dn <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        listG3_dn <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
        listG4_dn <- data4T[data4T[paste("logFC",contrast.l[4],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[4],sep=".")] < padj,]
        
      }
      
      ##### up ######
      cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
      cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]
      #Creem l'objecte del Venn
      if (symbols){
        listG1_up<-listG1_up[!is.na(listG1_up[,colnmes[2]]),]
        listG2_up<-listG2_up[!is.na(listG2_up[,colnmes[2]]),]
        listG3_up<-listG3_up[!is.na(listG3_up[,colnmes[2]]),]
        listG4_up<-listG4_up[!is.na(listG4_up[,colnmes[2]]),]
        
        list1<-listG1_up[,colnmes[2]]
        list2<-listG2_up[,colnmes[2]]
        list3<-listG3_up[,colnmes[2]]
        list4<-listG4_up[,colnmes[2]]
        
      }else{
        list1<-listG1_up[,colnmes[1]]
        list2<-listG2_up[,colnmes[1]]
        list3<-listG3_up[,colnmes[1]]
        list4<-listG4_up[,colnmes[1]]
        
      }
      
      #Creem l'objecte del Venn
      list.venn<-list(list1,list2,list3, list4)
      names(list.venn)<-c(listNames[1],listNames[2],listNames[3], listNames[4])
      vtest<-Venn(list.venn)

      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
      if(symbols){
          
          if(pval!=0){
            
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),] 
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] > logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 1 
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3]),] 
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] > logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 2
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4]),]
            data3.f <- data3[data3[paste("logFC",contrast.l[2],sep=".")] > logFC & data3[paste("P.Value", contrast.l[2],sep=".")] < pval &
                               data3[paste("logFC",contrast.l[1],sep=".")] > logFC & data3[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 2 + 1
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5]),]
            data4.f <- data4[data4[paste("logFC",contrast.l[3],sep=".")] > logFC & data4[paste("P.Value", contrast.l[3],sep=".")] < pval,] # contrast 3
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6]),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] > logFC & data5[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               data5[paste("logFC",contrast.l[1],sep=".")] > logFC & data5[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 1 + 3
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7]),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] > logFC & data6[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               data6[paste("logFC",contrast.l[2],sep=".")] > logFC & data6[paste("P.Value", contrast.l[2],sep=".")] < pval,] #contrast 2 + 3
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8]),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] > logFC & data7[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               data7[paste("logFC",contrast.l[2],sep=".")] > logFC & data7[paste("P.Value", contrast.l[2],sep=".")] < pval &
                               data7[paste("logFC",contrast.l[1],sep=".")] > logFC & data7[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 2 + 3 + 1
            data8 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9]),]
            data8.f <- data8[data8[paste("logFC",contrast.l[4],sep=".")] > logFC & data8[paste("P.Value", contrast.l[4],sep=".")] < pval,] # contrast 4
            data9 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10]),]
            data9.f <- data9[data9[paste("logFC",contrast.l[4],sep=".")] > logFC & data9[paste("P.Value", contrast.l[4],sep=".")] < pval &
                               data9[paste("logFC",contrast.l[1],sep=".")] > logFC & data9[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 +1
            data10 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11]),]
            data10.f <- data10[data10[paste("logFC",contrast.l[4],sep=".")] > logFC & data10[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data10[paste("logFC",contrast.l[2],sep=".")] > logFC & data10[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 2
            data11 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12]),]
            data11.f <- data11[data11[paste("logFC",contrast.l[4],sep=".")] > logFC & data11[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data11[paste("logFC",contrast.l[1],sep=".")] > logFC & data11[paste("P.Value", contrast.l[1],sep=".")] < pval &
                                 data11[paste("logFC",contrast.l[2],sep=".")] > logFC & data11[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 1 + 2
            data12 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13]),]
            data12.f <- data12[data12[paste("logFC",contrast.l[4],sep=".")] > logFC & data12[paste("P.Value", contrast.l[4],sep=".")] < pval & 
                                 data12[paste("logFC",contrast.l[3],sep=".")] > logFC & data12[paste("P.Value", contrast.l[3],sep=".")] < pval,] # contrast 4 + 3
            data13 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14]),]
            data13.f <- data13[data13[paste("logFC",contrast.l[4],sep=".")] > logFC & data13[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data13[paste("logFC",contrast.l[3],sep=".")] > logFC & data13[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 data13[paste("logFC",contrast.l[1],sep=".")] > logFC & data13[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 + 3 +1
            data14 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15]),]
            data14.f <- data14[data14[paste("logFC",contrast.l[4],sep=".")] > logFC & data14[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data14[paste("logFC",contrast.l[3],sep=".")] > logFC & data14[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 data14[paste("logFC",contrast.l[2],sep=".")] > logFC & data14[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 3 +2
            data15 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16]),]
            data15.f <- data15[data15[paste("logFC",contrast.l[4],sep=".")] > logFC & data15[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data15[paste("logFC",contrast.l[3],sep=".")] > logFC & data15[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 data15[paste("logFC",contrast.l[2],sep=".")] > logFC & data15[paste("P.Value", contrast.l[2],sep=".")] < pval &
                                 data15[paste("logFC",contrast.l[1],sep=".")] > logFC & data15[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 + 3 +2 +1
          }else{
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),]
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] > logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 1
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3]),]
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] > logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 2
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4]),]
            data3.f <- data3[data3[paste("logFC",contrast.l[2],sep=".")] > logFC & data3[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                               data3[paste("logFC",contrast.l[1],sep=".")] > logFC & data3[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 2 + 1
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5]),]
            data4.f <- data4[data4[paste("logFC",contrast.l[3],sep=".")] > logFC & data4[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]  # contrast 3
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6]),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] > logFC & data5[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & 
                               data5[paste("logFC",contrast.l[1],sep=".")] > logFC & data5[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 1 + 3
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7]),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] > logFC & data6[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               data6[paste("logFC",contrast.l[2],sep=".")] > logFC & data6[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] #contrast 2 + 3
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8]),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] > logFC & data7[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               data7[paste("logFC",contrast.l[2],sep=".")] > logFC & data7[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                               data7[paste("logFC",contrast.l[1],sep=".")] > logFC & data7[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 2 + 3 + 1
            data8 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9]),]
            data8.f <- data8[data8[paste("logFC",contrast.l[4],sep=".")] > logFC & data8[paste("adj.P.Val", contrast.l[4],sep=".")] < padj,] # contrast 4
            data9 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10]),]
            data9.f <- data9[data9[paste("logFC",contrast.l[4],sep=".")] > logFC & data9[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                               data9[paste("logFC",contrast.l[1],sep=".")] > logFC & data9[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 +1
            data10 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11]),]
            data10.f <- data10[data10[paste("logFC",contrast.l[4],sep=".")] > logFC & data10[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data10[paste("logFC",contrast.l[2],sep=".")] > logFC & data10[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 2
            data11 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12]),]
            data11.f <- data11[data11[paste("logFC",contrast.l[4],sep=".")] > logFC & data11[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data11[paste("logFC",contrast.l[2],sep=".")] > logFC & data11[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                                 data11[paste("logFC",contrast.l[1],sep=".")] > logFC & data11[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 1 + 2
            data12 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13]),]
            data12.f <- data12[data12[paste("logFC",contrast.l[4],sep=".")] > logFC & data12[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data12[paste("logFC",contrast.l[3],sep=".")] > logFC & data12[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,] # contrast 4 + 3
            data13 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14]),]
            data13.f <- data13[data13[paste("logFC",contrast.l[4],sep=".")] > logFC & data13[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data13[paste("logFC",contrast.l[3],sep=".")] > logFC & data13[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 data13[paste("logFC",contrast.l[1],sep=".")] > logFC & data13[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 3 +1
            data14 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15]),]
            data14.f <- data14[data14[paste("logFC",contrast.l[4],sep=".")] > logFC & data14[paste("adj.P.Val", contrast.l[4],sep=".")] < padj & 
                                 data14[paste("logFC",contrast.l[3],sep=".")] > logFC & data14[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 data14[paste("logFC",contrast.l[2],sep=".")] > logFC & data14[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 3 +2
            data15 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16]),]
            data15.f <- data15[data15[paste("logFC",contrast.l[4],sep=".")] > logFC & data15[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data15[paste("logFC",contrast.l[3],sep=".")] > logFC & data15[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 data15[paste("logFC",contrast.l[2],sep=".")] > logFC & data15[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                                 data15[paste("logFC",contrast.l[1],sep=".")] > logFC & data15[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 3 +2 +1
          }
          
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data1.f[(!is.na(data1.f[,colnmes[2]])) & 
                                 (data1.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]), cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue2")
          writeData(wb,data2.f[(!is.na(data2.f[,colnmes[2]]))  & 
                                 (data2.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]), cNames], 
                    sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data3.f[(!is.na(data3.f[,colnmes[2]])) & 
                                 ((data3.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) | 
                                    (data3.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])), cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown1")
          writeData(wb,data4.f[(!is.na(data4.f[,colnmes[2]]))  & 
                                 (data4.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]), cNames], 
                    sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange1")
          writeData(wb,data5.f[(!is.na(data5.f[,colnmes[2]]))  & 
                                 ((data5.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                    (data5.f[,colnmes[1]] %in% listG1_up[,colnmes[1]])), cNames], 
                    sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow1")
          writeData(wb,data6.f[(!is.na(data6.f[,colnmes[2]])) & 
                                 ((data6.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                    (data6.f[,colnmes[1]] %in% listG2_up[,colnmes[1]])), cNames], 
                    sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown2")
          writeData(wb,data7.f[(!is.na(data7.f[,colnmes[2]]))  & 
                                 ((data7.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG2_up[,colnmes[1]])), cNames], 
                    sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "red")
          writeData(wb,data8.f[(!is.na(data8.f[,colnmes[2]]))  & 
                                 (data8.f[,colnmes[1]] %in% listG4_up[,colnmes[1]]), cNames], 
                    sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green3")
          writeData(wb,data9.f[(!is.na(data9.f[,colnmes[2]]))  & 
                                 ((data9.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) | 
                                    (data9.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange2")
          writeData(wb,data10.f[(!is.na(data10.f[,colnmes[2]]))  & 
                                  ((data10.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]) | 
                                     (data10.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue1")
          writeData(wb,data11.f[(!is.na(data11.f[,colnmes[2]]))  & 
                                  ((data11.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) | 
                                     (data11.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]) |
                                     (data11.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple1")
          writeData(wb,data12.f[(!is.na(data12.f[,colnmes[2]]))  & 
                                  ((data12.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                     (data12.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data13.f[(!is.na(data13.f[,colnmes[2]]))  & 
                                  ((data13.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) |
                                     (data13.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                     (data13.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow2")
          writeData(wb,data14.f[(!is.na(data14.f[,colnmes[2]]))  & 
                                  ((data14.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]) |
                                     (data14.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                     (data14.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple2")
          writeData(wb,data15.f[(!is.na(data15.f[,colnmes[2]]))  & 
                                  ((data15.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG4_up[,colnmes[1]])),cNames], 
                    sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
          
          saveWorkbook(wb,file=paste(resultsDir,
                                     paste("UP.VennGenes",filename,"xlsx",sep="."),
                                     sep="/"),overwrite = TRUE)
          vennData <- c(length(unique(data1.f$Symbol)),length(unique(data2.f$Symbol)),length(unique(data3.f$Symbol)),
                        length(unique(data4.f$Symbol)),length(unique(data5.f$Symbol)),length(unique(data6.f$Symbol)),
                        length(unique(data7.f$Symbol)),length(unique(data8.f$Symbol)),length(unique(data9.f$Symbol)),
                        length(unique(data10.f$Symbol)),length(unique(data11.f$Symbol)),length(unique(data12.f$Symbol)),
                        length(unique(data13.f$Symbol)),length(unique(data14.f$Symbol)),length(unique(data15.f$Symbol)))
          
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("UP.VennDiagram",filename,"png",sep=".")),width = 1200, height = 700)
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("UP.VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 700)
          }
          plotVenn4d(vennData, labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
          
        }else{
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[2]),cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[3]),cNames], 
                    sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[4]),cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[5]),cNames], 
                    sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[6]),cNames], 
                    sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[7]),cNames], 
                    sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[8]),cNames], 
                    sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "red")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[9]),cNames], 
                    sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green3")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[10]),cNames], 
                    sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[11]),cNames], 
                    sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[12]),cNames], 
                    sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[13]),cNames], 
                    sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[14]),cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[15]),cNames], 
                    sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[16]),cNames], 
                    sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
          
          saveWorkbook(wb,file=paste(resultsDir,
                                     paste("UP.VennGenes",filename,"xlsx",sep="."),
                                     sep="/"),overwrite = TRUE)
          
          vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
          
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("UP.VennDiagram",filename,"png",sep=".")),width = 1200, height = 700)
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("UP.VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 700)
          }
          plotVenn4d(vennData[-1], labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }
      ##### down ######
      cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
      cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]
      #Creem l'objecte del Venn
      if (symbols){
        listG1_dn<-listG1_dn[!is.na(listG1_dn[,colnmes[2]]),]
        listG2_dn<-listG2_dn[!is.na(listG2_dn[,colnmes[2]]),]
        listG3_dn<-listG3_dn[!is.na(listG3_dn[,colnmes[2]]),]
        listG4_dn<-listG4_dn[!is.na(listG4_dn[,colnmes[2]]),]
        
        list1<-listG1_dn[,colnmes[2]]
        list2<-listG2_dn[,colnmes[2]]
        list3<-listG3_dn[,colnmes[2]]
        list4<-listG4_dn[,colnmes[2]]
        
      }else{
        list1<-listG1_dn[,colnmes[1]]
        list2<-listG2_dn[,colnmes[1]]
        list3<-listG3_dn[,colnmes[1]]
        list4<-listG4_dn[,colnmes[1]]
        
      }
      
      #Creem l'objecte del Venn
      list.venn<-list(list1,list2,list3, list4)
      names(list.venn)<-c(listNames[1],listNames[2],listNames[3], listNames[4])
      vtest<-Venn(list.venn)
      
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
      if(symbols){
          
          if(pval!=0){
            
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),] 
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] < -logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 1 
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3]),] 
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] <- logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 2
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4]),]
            data3.f <- data3[data3[paste("logFC",contrast.l[2],sep=".")] < -logFC & data3[paste("P.Value", contrast.l[2],sep=".")] < pval &
                               data3[paste("logFC",contrast.l[1],sep=".")] < -logFC & data3[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 2 + 1
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5]),]
            data4.f <- data4[data4[paste("logFC",contrast.l[3],sep=".")] < -logFC & data4[paste("P.Value", contrast.l[3],sep=".")] < pval,] # contrast 3
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6]),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] < -logFC & data5[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               data5[paste("logFC",contrast.l[1],sep=".")] < -logFC & data5[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 1 + 3
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7]),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] < -logFC & data6[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               data6[paste("logFC",contrast.l[2],sep=".")] < -logFC & data6[paste("P.Value", contrast.l[2],sep=".")] < pval,] #contrast 2 + 3
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8]),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] < -logFC & data7[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               data7[paste("logFC",contrast.l[2],sep=".")] < -logFC & data7[paste("P.Value", contrast.l[2],sep=".")] < pval &
                               data7[paste("logFC",contrast.l[1],sep=".")] < -logFC & data7[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 2 + 3 + 1
            data8 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9]),]
            data8.f <- data8[data8[paste("logFC",contrast.l[4],sep=".")] < -logFC & data8[paste("P.Value", contrast.l[4],sep=".")] < pval,] # contrast 4
            data9 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10]),]
            data9.f <- data9[data9[paste("logFC",contrast.l[4],sep=".")] < -logFC & data9[paste("P.Value", contrast.l[4],sep=".")] < pval &
                               data9[paste("logFC",contrast.l[1],sep=".")] < -logFC & data9[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 +1
            data10 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11]),]
            data10.f <- data10[data10[paste("logFC",contrast.l[4],sep=".")] < -logFC & data10[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data10[paste("logFC",contrast.l[2],sep=".")] < -logFC & data10[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 2
            data11 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12]),]
            data11.f <- data11[data11[paste("logFC",contrast.l[4],sep=".")] < -logFC & data11[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data11[paste("logFC",contrast.l[1],sep=".")] < -logFC & data11[paste("P.Value", contrast.l[1],sep=".")] < pval &
                                 data11[paste("logFC",contrast.l[2],sep=".")] < -logFC & data11[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 1 + 2
            data12 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13]),]
            data12.f <- data12[data12[paste("logFC",contrast.l[4],sep=".")] < -logFC & data12[paste("P.Value", contrast.l[4],sep=".")] < pval & 
                                 data12[paste("logFC",contrast.l[3],sep=".")] < -logFC & data12[paste("P.Value", contrast.l[3],sep=".")] < pval,] # contrast 4 + 3
            data13 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14]),]
            data13.f <- data13[data13[paste("logFC",contrast.l[4],sep=".")] < -logFC & data13[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data13[paste("logFC",contrast.l[3],sep=".")] < -logFC & data13[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 data13[paste("logFC",contrast.l[1],sep=".")] < -logFC & data13[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 + 3 +1
            data14 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15]),]
            data14.f <- data14[data14[paste("logFC",contrast.l[4],sep=".")] < -logFC & data14[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data14[paste("logFC",contrast.l[3],sep=".")] < -logFC & data14[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 data14[paste("logFC",contrast.l[2],sep=".")] < -logFC & data14[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 3 +2
            data15 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16]),]
            data15.f <- data15[data15[paste("logFC",contrast.l[4],sep=".")] < -logFC & data15[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 data15[paste("logFC",contrast.l[3],sep=".")] < -logFC & data15[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 data15[paste("logFC",contrast.l[2],sep=".")] < -logFC & data15[paste("P.Value", contrast.l[2],sep=".")] < pval &
                                 data15[paste("logFC",contrast.l[1],sep=".")] < -logFC & data15[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 + 3 +2 +1
          }else{
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),]
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] < -logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 1
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3]),]
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] < -logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 2
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4]),]
            data3.f <- data3[data3[paste("logFC",contrast.l[2],sep=".")] < -logFC & data3[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                               data3[paste("logFC",contrast.l[1],sep=".")] < -logFC & data3[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 2 + 1
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5]),]
            data4.f <- data4[data4[paste("logFC",contrast.l[3],sep=".")] < -logFC & data4[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]  # contrast 3
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6]),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] < -logFC & data5[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & 
                               data5[paste("logFC",contrast.l[1],sep=".")] < -logFC & data5[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 1 + 3
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7]),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] < -logFC & data6[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               data6[paste("logFC",contrast.l[2],sep=".")] < -logFC & data6[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] #contrast 2 + 3
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8]),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] < -logFC & data7[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               data7[paste("logFC",contrast.l[2],sep=".")] < -logFC & data7[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                               data7[paste("logFC",contrast.l[1],sep=".")] < -logFC & data7[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 2 + 3 + 1
            data8 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9]),]
            data8.f <- data8[data8[paste("logFC",contrast.l[4],sep=".")] < -logFC & data8[paste("adj.P.Val", contrast.l[4],sep=".")] < padj,] # contrast 4
            data9 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10]),]
            data9.f <- data9[data9[paste("logFC",contrast.l[4],sep=".")] < -logFC & data9[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                               data9[paste("logFC",contrast.l[1],sep=".")] < -logFC & data9[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 +1
            data10 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11]),]
            data10.f <- data10[data10[paste("logFC",contrast.l[4],sep=".")] < -logFC & data10[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data10[paste("logFC",contrast.l[2],sep=".")] < -logFC & data10[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 2
            data11 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12]),]
            data11.f <- data11[data11[paste("logFC",contrast.l[4],sep=".")] < -logFC & data11[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data11[paste("logFC",contrast.l[2],sep=".")] < -logFC & data11[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                                 data11[paste("logFC",contrast.l[1],sep=".")] < -logFC & data11[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 1 + 2
            data12 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13]),]
            data12.f <- data12[data12[paste("logFC",contrast.l[4],sep=".")] < -logFC & data12[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data12[paste("logFC",contrast.l[3],sep=".")] < -logFC & data12[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,] # contrast 4 + 3
            data13 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14]),]
            data13.f <- data13[data13[paste("logFC",contrast.l[4],sep=".")] < -logFC & data13[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data13[paste("logFC",contrast.l[3],sep=".")] < -logFC & data13[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 data13[paste("logFC",contrast.l[1],sep=".")] < -logFC & data13[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 3 +1
            data14 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15]),]
            data14.f <- data14[data14[paste("logFC",contrast.l[4],sep=".")] < -logFC & data14[paste("adj.P.Val", contrast.l[4],sep=".")] < padj & 
                                 data14[paste("logFC",contrast.l[3],sep=".")] < -logFC & data14[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 data14[paste("logFC",contrast.l[2],sep=".")] < -logFC & data14[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 3 +2
            data15 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16]),]
            data15.f <- data15[data15[paste("logFC",contrast.l[4],sep=".")] < -logFC & data15[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 data15[paste("logFC",contrast.l[3],sep=".")] < -logFC & data15[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 data15[paste("logFC",contrast.l[2],sep=".")] < -logFC & data15[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                                 data15[paste("logFC",contrast.l[1],sep=".")] < -logFC & data15[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 3 +2 +1
          }
          
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data1.f[(!is.na(data1.f[,colnmes[2]])) & 
                                 (data1.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]), cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue2")
          writeData(wb,data2.f[(!is.na(data2.f[,colnmes[2]]))  & 
                                 (data2.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]), cNames], 
                    sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data3.f[(!is.na(data3.f[,colnmes[2]])) & 
                                 ((data3.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) | 
                                    (data3.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])), cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown1")
          writeData(wb,data4.f[(!is.na(data4.f[,colnmes[2]]))  & 
                                 (data4.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]), cNames], 
                    sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange1")
          writeData(wb,data5.f[(!is.na(data5.f[,colnmes[2]]))  & 
                                 ((data5.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                    (data5.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]])), cNames], 
                    sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow1")
          writeData(wb,data6.f[(!is.na(data6.f[,colnmes[2]])) & 
                                 ((data6.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                    (data6.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]])), cNames], 
                    sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown2")
          writeData(wb,data7.f[(!is.na(data7.f[,colnmes[2]]))  & 
                                 ((data7.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]])), cNames], 
                    sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "red")
          writeData(wb,data8.f[(!is.na(data8.f[,colnmes[2]]))  & 
                                 (data8.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]]), cNames], 
                    sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green3")
          writeData(wb,data9.f[(!is.na(data9.f[,colnmes[2]]))  & 
                                 ((data9.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) | 
                                    (data9.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange2")
          writeData(wb,data10.f[(!is.na(data10.f[,colnmes[2]]))  & 
                                  ((data10.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]) | 
                                     (data10.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue1")
          writeData(wb,data11.f[(!is.na(data11.f[,colnmes[2]]))  & 
                                  ((data11.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) | 
                                     (data11.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]) |
                                     (data11.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple1")
          writeData(wb,data12.f[(!is.na(data12.f[,colnmes[2]]))  & 
                                  ((data12.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                     (data12.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data13.f[(!is.na(data13.f[,colnmes[2]]))  & 
                                  ((data13.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) |
                                     (data13.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                     (data13.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow2")
          writeData(wb,data14.f[(!is.na(data14.f[,colnmes[2]]))  & 
                                  ((data14.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]) |
                                     (data14.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                     (data14.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple2")
          writeData(wb,data15.f[(!is.na(data15.f[,colnmes[2]]))  & 
                                  ((data15.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG4_dn[,colnmes[1]])),cNames], 
                    sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
          
          saveWorkbook(wb,file=paste(resultsDir,
                                     paste("DN.VennGenes",filename,"xlsx",sep="."),
                                     sep="/"),overwrite = TRUE)
          vennData <- c(length(unique(data1.f$Symbol)),length(unique(data2.f$Symbol)),length(unique(data3.f$Symbol)),
                        length(unique(data4.f$Symbol)),length(unique(data5.f$Symbol)),length(unique(data6.f$Symbol)),
                        length(unique(data7.f$Symbol)),length(unique(data8.f$Symbol)),length(unique(data9.f$Symbol)),
                        length(unique(data10.f$Symbol)),length(unique(data11.f$Symbol)),length(unique(data12.f$Symbol)),
                        length(unique(data13.f$Symbol)),length(unique(data14.f$Symbol)),length(unique(data15.f$Symbol)))
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("DN.VennDiagram",filename,"png",sep=".")),width = 1200, height = 700)
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("DN.VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 700)
          }
          plotVenn4d(vennData, labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }else{
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[2]),cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[3]),cNames], 
                    sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[4]),cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[5]),cNames], 
                    sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[6]),cNames], 
                    sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[7]),cNames], 
                    sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[8]),cNames], 
                    sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "red")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[9]),cNames], 
                    sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green3")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[10]),cNames], 
                    sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[11]),cNames], 
                    sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[12]),cNames], 
                    sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[13]),cNames], 
                    sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[14]),cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[15]),cNames], 
                    sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[16]),cNames], 
                    sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
          
          saveWorkbook(wb,file=paste(resultsDir,
                                     paste("DN.VennGenes",filename,"xlsx",sep="."),
                                     sep="/"),overwrite = TRUE)
          
          vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("DN.VennDiagram",filename,"png",sep=".")),width = 1200, height = 700)
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("DN.VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 700)
          }
          plotVenn4d(vennData[-1], labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }
    }else{
      if(pval!=0){
        listG1 <- data4T[abs(data4T[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2 <- data4T[abs(data4T[paste("logFC",contrast.l[2],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        listG3 <- data4T[abs(data4T[paste("logFC",contrast.l[3],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[3],sep=".")] < pval,]
        listG4 <- data4T[abs(data4T[paste("logFC",contrast.l[4],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[4],sep=".")] < pval,]
        
      }else{
        listG1 <- data4T[abs(data4T[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2 <- data4T[abs(data4T[paste("logFC",contrast.l[2],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        listG3 <- data4T[abs(data4T[paste("logFC",contrast.l[3],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
        listG4 <- data4T[abs(data4T[paste("logFC",contrast.l[4],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[4],sep=".")] < padj,]
      }
      
      cols <- c(brewer.pal(8,"Pastel1"), brewer.pal(8,"Pastel2"))  #Fixar els colors per als venn diagrams amb 4 condicions
      cols <- cols[c(8,2,3,15,5,6,7,1,9,10,11,12,13,14,4,16)]
      
      #Creem l'objecte del Venn
      if(symbols) {
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
      
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
      if(symbols){
          
          if(pval!=0){
            
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),] 
            data1.f <- data1[abs(data1[paste("logFC",contrast.l[1],sep=".")]) > logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 1 
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3]),] 
            data2.f <- data2[abs(data2[paste("logFC",contrast.l[2],sep=".")]) > logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 2
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4]),]
            data3.f <- data3[abs(data3[paste("logFC",contrast.l[2],sep=".")]) > logFC & data3[paste("P.Value", contrast.l[2],sep=".")] < pval &
                               abs(data3[paste("logFC",contrast.l[1],sep=".")]) > logFC & data3[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 2 + 1
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5]),]
            data4.f <- data4[abs(data4[paste("logFC",contrast.l[3],sep=".")]) > logFC & data4[paste("P.Value", contrast.l[3],sep=".")] < pval,] # contrast 3
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6]),]
            data5.f <- data5[abs(data5[paste("logFC",contrast.l[3],sep=".")]) > logFC & data5[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               abs(data5[paste("logFC",contrast.l[1],sep=".")]) > logFC & data5[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 1 + 3
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7]),]
            data6.f <- data6[abs(data6[paste("logFC",contrast.l[3],sep=".")]) > logFC & data6[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               abs(data6[paste("logFC",contrast.l[2],sep=".")]) > logFC & data6[paste("P.Value", contrast.l[2],sep=".")] < pval,] #contrast 2 + 3
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8]),]
            data7.f <- data7[abs(data7[paste("logFC",contrast.l[3],sep=".")]) > logFC & data7[paste("P.Value", contrast.l[3],sep=".")] < pval &
                               abs(data7[paste("logFC",contrast.l[2],sep=".")]) > logFC & data7[paste("P.Value", contrast.l[2],sep=".")] < pval &
                               abs(data7[paste("logFC",contrast.l[1],sep=".")]) > logFC & data7[paste("P.Value", contrast.l[1],sep=".")] < pval,] #contrast 2 + 3 + 1
            data8 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9]),]
            data8.f <- data8[abs(data8[paste("logFC",contrast.l[4],sep=".")]) > logFC & data8[paste("P.Value", contrast.l[4],sep=".")] < pval,] # contrast 4
            data9 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10]),]
            data9.f <- data9[abs(data9[paste("logFC",contrast.l[4],sep=".")]) > logFC & data9[paste("P.Value", contrast.l[4],sep=".")] < pval &
                               abs(data9[paste("logFC",contrast.l[1],sep=".")]) > logFC & data9[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 +1
            data10 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11]),]
            data10.f <- data10[abs(data10[paste("logFC",contrast.l[4],sep=".")]) > logFC & data10[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 abs(data10[paste("logFC",contrast.l[2],sep=".")]) > logFC & data10[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 2
            data11 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12]),]
            data11.f <- data11[abs(data11[paste("logFC",contrast.l[4],sep=".")]) > logFC & data11[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 abs(data11[paste("logFC",contrast.l[1],sep=".")]) > logFC & data11[paste("P.Value", contrast.l[1],sep=".")] < pval &
                                 abs(data11[paste("logFC",contrast.l[2],sep=".")]) > logFC & data11[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 1 + 2
            data12 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13]),]
            data12.f <- data12[abs(data12[paste("logFC",contrast.l[4],sep=".")]) > logFC & data12[paste("P.Value", contrast.l[4],sep=".")] < pval & 
                                 abs(data12[paste("logFC",contrast.l[3],sep=".")]) > logFC & data12[paste("P.Value", contrast.l[3],sep=".")] < pval,] # contrast 4 + 3
            data13 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14]),]
            data13.f <- data13[abs(data13[paste("logFC",contrast.l[4],sep=".")]) > logFC & data13[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 abs(data13[paste("logFC",contrast.l[3],sep=".")]) > logFC & data13[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 abs(data13[paste("logFC",contrast.l[1],sep=".")]) > logFC & data13[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 + 3 +1
            data14 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15]),]
            data14.f <- data14[abs(data14[paste("logFC",contrast.l[4],sep=".")]) > logFC & data14[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 abs(data14[paste("logFC",contrast.l[3],sep=".")]) > logFC & data14[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 abs(data14[paste("logFC",contrast.l[2],sep=".")]) > logFC & data14[paste("P.Value", contrast.l[2],sep=".")] < pval,] # contrast 4 + 3 +2
            data15 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16]),]
            data15.f <- data15[abs(data15[paste("logFC",contrast.l[4],sep=".")]) > logFC & data15[paste("P.Value", contrast.l[4],sep=".")] < pval &
                                 abs(data15[paste("logFC",contrast.l[3],sep=".")]) > logFC & data15[paste("P.Value", contrast.l[3],sep=".")] < pval &
                                 abs(data15[paste("logFC",contrast.l[2],sep=".")]) > logFC & data15[paste("P.Value", contrast.l[2],sep=".")] < pval &
                                 abs(data15[paste("logFC",contrast.l[1],sep=".")]) > logFC & data15[paste("P.Value", contrast.l[1],sep=".")] < pval,] # contrast 4 + 3 +2 +1
          }else{
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),]
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[2]),] 
            data1.f <- data1[abs(data1[paste("logFC",contrast.l[1],sep=".")]) > logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 1 
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[3]),] 
            data2.f <- data2[abs(data2[paste("logFC",contrast.l[2],sep=".")]) > logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 2
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[4]),]
            data3.f <- data3[abs(data3[paste("logFC",contrast.l[2],sep=".")]) > logFC & data3[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                               abs(data3[paste("logFC",contrast.l[1],sep=".")]) > logFC & data3[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 2 + 1
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[5]),]
            data4.f <- data4[abs(data4[paste("logFC",contrast.l[3],sep=".")]) > logFC & data4[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,] # contrast 3
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[6]),]
            data5.f <- data5[abs(data5[paste("logFC",contrast.l[3],sep=".")]) > logFC & data5[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               abs(data5[paste("logFC",contrast.l[1],sep=".")]) > logFC & data5[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 1 + 3
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[7]),]
            data6.f <- data6[abs(data6[paste("logFC",contrast.l[3],sep=".")]) > logFC & data6[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               abs(data6[paste("logFC",contrast.l[2],sep=".")]) > logFC & data6[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] #contrast 2 + 3
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[8]),]
            data7.f <- data7[abs(data7[paste("logFC",contrast.l[3],sep=".")]) > logFC & data7[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                               abs(data7[paste("logFC",contrast.l[2],sep=".")]) > logFC & data7[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                               abs(data7[paste("logFC",contrast.l[1],sep=".")]) > logFC & data7[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] #contrast 2 + 3 + 1
            data8 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[9]),]
            data8.f <- data8[abs(data8[paste("logFC",contrast.l[4],sep=".")]) > logFC & data8[paste("adj.P.Val", contrast.l[4],sep=".")] < padj,] # contrast 4
            data9 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[10]),]
            data9.f <- data9[abs(data9[paste("logFC",contrast.l[4],sep=".")]) > logFC & data9[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                               abs(data9[paste("logFC",contrast.l[1],sep=".")]) > logFC & data9[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 +1
            data10 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[11]),]
            data10.f <- data10[abs(data10[paste("logFC",contrast.l[4],sep=".")]) > logFC & data10[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 abs(data10[paste("logFC",contrast.l[2],sep=".")]) > logFC & data10[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 2
            data11 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[12]),]
            data11.f <- data11[abs(data11[paste("logFC",contrast.l[4],sep=".")]) > logFC & data11[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 abs(data11[paste("logFC",contrast.l[1],sep=".")]) > logFC & data11[paste("adj.P.Val", contrast.l[1],sep=".")] < padj &
                                 abs(data11[paste("logFC",contrast.l[2],sep=".")]) > logFC & data11[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 1 + 2
            data12 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[13]),]
            data12.f <- data12[abs(data12[paste("logFC",contrast.l[4],sep=".")]) > logFC & data12[paste("adj.P.Val", contrast.l[4],sep=".")] < padj & 
                                 abs(data12[paste("logFC",contrast.l[3],sep=".")]) > logFC & data12[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,] # contrast 4 + 3
            data13 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[14]),]
            data13.f <- data13[abs(data13[paste("logFC",contrast.l[4],sep=".")]) > logFC & data13[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 abs(data13[paste("logFC",contrast.l[3],sep=".")]) > logFC & data13[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 abs(data13[paste("logFC",contrast.l[1],sep=".")]) > logFC & data13[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,] # contrast 4 + 3 +1
            data14 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[15]),]
            data14.f <- data14[abs(data14[paste("logFC",contrast.l[4],sep=".")]) > logFC & data14[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 abs(data14[paste("logFC",contrast.l[3],sep=".")]) > logFC & data14[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 abs(data14[paste("logFC",contrast.l[2],sep=".")]) > logFC & data14[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,] # contrast 4 + 3 +2
            data15 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets[16]),]
            data15.f <- data15[abs(data15[paste("logFC",contrast.l[4],sep=".")]) > logFC & data15[paste("adj.P.Val", contrast.l[4],sep=".")] < padj &
                                 abs(data15[paste("logFC",contrast.l[3],sep=".")]) > logFC & data15[paste("adj.P.Val", contrast.l[3],sep=".")] < padj &
                                 abs(data15[paste("logFC",contrast.l[2],sep=".")]) > logFC & data15[paste("adj.P.Val", contrast.l[2],sep=".")] < padj &
                                 abs(data15[paste("logFC",contrast.l[1],sep=".")]) > logFC & data15[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          }
        
        
          
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data1.f[(!is.na(data1.f[,colnmes[2]])) & 
                                 (data1.f[,colnmes[1]] %in% listG1[,colnmes[1]]), cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue2")
          writeData(wb,data2.f[(!is.na(data2.f[,colnmes[2]]))  & 
                                 (data2.f[,colnmes[1]] %in% listG2[,colnmes[1]]), cNames], 
                    sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data3.f[(!is.na(data3.f[,colnmes[2]])) & 
                                 ((data3.f[,colnmes[1]] %in% listG1[,colnmes[1]]) | 
                                    (data3.f[,colnmes[1]] %in% listG4[,colnmes[1]])), cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown1")
          writeData(wb,data4.f[(!is.na(data4.f[,colnmes[2]]))  & 
                                 (data4.f[,colnmes[1]] %in% listG3[,colnmes[1]]), cNames], 
                    sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange1")
          writeData(wb,data5.f[(!is.na(data5.f[,colnmes[2]]))  & 
                                 ((data5.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data5.f[,colnmes[1]] %in% listG1[,colnmes[1]])), cNames], 
                    sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow1")
          writeData(wb,data6.f[(!is.na(data6.f[,colnmes[2]])) & 
                                 ((data6.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data6.f[,colnmes[1]] %in% listG2[,colnmes[1]])), cNames], 
                    sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown2")
          writeData(wb,data7.f[(!is.na(data7.f[,colnmes[2]]))  & 
                                 ((data7.f[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG2[,colnmes[1]])), cNames], 
                    sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "red")
          writeData(wb,data8.f[(!is.na(data8.f[,colnmes[2]]))  & 
                                 (data8.f[,colnmes[1]] %in% listG4[,colnmes[1]]), cNames], 
                    sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green3")
          writeData(wb,data9.f[(!is.na(data9.f[,colnmes[2]]))  & 
                                 ((data9.f[,colnmes[1]] %in% listG1[,colnmes[1]]) | 
                                    (data9.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange2")
          writeData(wb,data10.f[(!is.na(data10.f[,colnmes[2]]))  & 
                                  ((data10.f[,colnmes[1]] %in% listG2[,colnmes[1]]) | 
                                     (data10.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue1")
          writeData(wb,data11.f[(!is.na(data11.f[,colnmes[2]]))  & 
                                  ((data11.f[,colnmes[1]] %in% listG1[,colnmes[1]]) | 
                                     (data11.f[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                     (data11.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple1")
          writeData(wb,data12.f[(!is.na(data12.f[,colnmes[2]]))  & 
                                  ((data12.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                     (data12.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data13.f[(!is.na(data13.f[,colnmes[2]]))  & 
                                  ((data13.f[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                     (data13.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                     (data13.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow2")
          writeData(wb,data14.f[(!is.na(data14.f[,colnmes[2]]))  & 
                                  ((data14.f[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                     (data14.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                     (data14.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple2")
          writeData(wb,data15.f[(!is.na(data15.f[,colnmes[2]]))  & 
                                  ((data15.f[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG3[,colnmes[1]]) |
                                     (data15.f[,colnmes[1]] %in% listG4[,colnmes[1]])),cNames], 
                    sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
          
          saveWorkbook(wb,file=paste(resultsDir,
                                     paste("VennGenes",filename,"xlsx",sep="."),
                                     sep="/"),overwrite = TRUE)
          
          vennData <- c(length(unique(data1.f$Symbol)),length(unique(data2.f$Symbol)),length(unique(data3.f$Symbol)),
                        length(unique(data4.f$Symbol)),length(unique(data5.f$Symbol)),length(unique(data6.f$Symbol)),
                        length(unique(data7.f$Symbol)),length(unique(data8.f$Symbol)),length(unique(data9.f$Symbol)),
                        length(unique(data10.f$Symbol)),length(unique(data11.f$Symbol)),length(unique(data12.f$Symbol)),
                        length(unique(data13.f$Symbol)),length(unique(data14.f$Symbol)),length(unique(data15.f$Symbol)))
          
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")),width = 1200, height = 600)
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 600)
          }
          plotVenn4d(vennData, labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
          
        }else{
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[2]),cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[3]),cNames], 
                    sheet = "blue2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[4]),cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[5]),cNames], 
                    sheet = "brown1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[6]),cNames], 
                    sheet = "orange1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[7]),cNames], 
                    sheet = "yellow1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[8]),cNames], 
                    sheet = "brown2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "red")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[9]),cNames], 
                    sheet = "red", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green3")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[10]),cNames], 
                    sheet = "green3", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[11]),cNames], 
                    sheet = "orange2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[12]),cNames], 
                    sheet = "blue1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[13]),cNames], 
                    sheet = "purple1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[14]),cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[15]),cNames], 
                    sheet = "yellow2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "purple2")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets[16]),cNames], 
                    sheet = "purple2", startRow = 1, startCol = 1, headerStyle = hs1)
          
          saveWorkbook(wb,file=paste(resultsDir,
                                     paste("VennGenes",filename,"xlsx",sep="."),
                                     sep="/"),overwrite = TRUE)
          vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))

          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")),width = 1200, height = 700)
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")),width = 1200, height = 700)
          }
          plotVenn4d(vennData[-1], labels=c(listNames[1],listNames[2],listNames[3], listNames[4]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }

      }
    }
  }
