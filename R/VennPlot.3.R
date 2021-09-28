VennPlot.3 <- function(contrast.l = NULL , listNames, filename, data4T= NULL, symbols=TRUE,
                       colnmes= c("AffyID", "Symbol"),listSymbols=F,listA=NULL, listB=NULL,listC=NULL,
                             img.fmt= "pdf",
                             pval=0, padj=0.05, logFC=1, up_down=FALSE){
  
  require(Vennerable) 
  require(colorfulVennPlot) #Per generar plots amb colors diferents
  require(RColorBrewer)  
  require(openxlsx)  
  
  if (listSymbols){
    listG1<-listA
    listG2<-listB
    listG3<-listC
    
    cols <- brewer.pal(8,"Pastel2") 
    
    listG1 <- listG1[!is.na(listG1)]
    listG2 <- listG2[!is.na(listG2)]
    listG3 <- listG3[!is.na(listG3)]
    
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
               Colors=cols, Title="", shrink=0.5)
    dev.off()
    if(!is.null(data4T)){
      #EXCEL
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
      addWorksheet(wb, sheetName = "CommonALL grey")
      writeData(wb,data4T[!is.na(data4T[,"Symbol"]) & (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`111`)),
                          cNames], 
                sheet = "CommonALL grey", startRow = 1, startCol = 1, headerStyle = hs1)
      saveWorkbook(wb,file=file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
      
    }else{
      genes <- data.frame("Genes"=unlist(list.venn))
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      
      wb <- createWorkbook()
      addWorksheet(wb, sheetName = "orange") 
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`100`),]), 
                sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "blue")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`010`),]), 
                sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "green")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`001`),]), 
                sheet = "green", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "pink")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`110`),]), 
                sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "brown")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`011`),]), 
                sheet = "brown", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "yellow")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`101`),]), 
                sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "CommonALL grey")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`111`),]), 
                sheet = "CommonALL grey", startRow = 1, startCol = 1, headerStyle = hs1)
      saveWorkbook(wb,file=file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
    }
  }else{
    if (up_down){
      if (pval!=0){
        listG1_up <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2_up <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        listG3_up <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[3],sep=".")] < pval,]
        
        
        listG1_dn <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2_dn <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        listG3_dn <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[3],sep=".")] < pval,]
        
      }else{
        listG1_up <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2_up <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        listG3_up <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
        
        
        listG1_dn <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2_dn <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        listG3_dn <- data4T[data4T[paste("logFC",contrast.l[3],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
      }
      ##### up ######
      cols <- brewer.pal(8,"Pastel2") 
      #Creem l'objecte del Venn
      if (symbols){
        listG1_up<-listG1_up[!is.na(listG1_up[,colnmes[2]]),]
        listG2_up<-listG2_up[!is.na(listG2_up[,colnmes[2]]),]
        listG3_up<-listG3_up[!is.na(listG3_up[,colnmes[2]]),]
        list1<-listG1_up[,colnmes[2]]
        list2<-listG2_up[,colnmes[2]]
        list3<-listG3_up[,colnmes[2]]
      }else{
        list1<-listG1_up[,colnmes[1]]
        list2<-listG2_up[,colnmes[1]]
        list3<-listG3_up[,colnmes[1]]
      }
      #Ens assegurem de que no hi ha NA
      list.venn<-list(list1,list2,list3)
      names(list.venn)<-c(listNames[1],listNames[2],listNames[3])
      vtest<-Venn(list.venn)
      #vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
      
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
        #cNames <- colnames(data4T)
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
      if (symbols) {
          
          if (pval!=0){
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`100`),]
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] > logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,]
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`010`),]
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] > logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`001`),]
            data3.f <-  data3[data3[paste("logFC",contrast.l[3],sep=".")] > logFC & data3[paste("P.Value", contrast.l[3],sep=".")] < pval,]
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`110`),]
            data4.f <- data4[data4[paste("logFC",contrast.l[1],sep=".")] > logFC & data4[paste("P.Value", contrast.l[1],sep=".")] < pval & data4[paste("logFC",contrast.l[2],sep=".")] > logFC &
                               data4[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`011`),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] > logFC & data5[paste("P.Value", contrast.l[3],sep=".")] < pval & data5[paste("logFC",contrast.l[2],sep=".")] > logFC &
                               data5[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`101`),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] > logFC & data6[paste("P.Value", contrast.l[3],sep=".")] < pval & data6[paste("logFC",contrast.l[1],sep=".")] > logFC &
                               data6[paste("P.Value", contrast.l[1],sep=".")] < pval,]
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`111`),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] > logFC & data7[paste("P.Value", contrast.l[3],sep=".")] < pval & data7[paste("logFC",contrast.l[2],sep=".")] > logFC &
                               data7[paste("P.Value", contrast.l[2],sep=".")] < pval & data7[paste("logFC",contrast.l[1],sep=".")] > logFC &
                               data7[paste("P.Value", contrast.l[1],sep=".")] < pval,]
          }else{
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`100`),]
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] > logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`010`),]
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] > logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`001`),]
            data3.f <-  data3[data3[paste("logFC",contrast.l[3],sep=".")] > logFC & data3[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`110`),]
            data4.f <- data4[data4[paste("logFC",contrast.l[1],sep=".")] > logFC & data4[paste("adj.P.Val", contrast.l[1],sep=".")] < padj & data4[paste("logFC",contrast.l[2],sep=".")] > logFC &
                               data4[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`011`),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] > logFC & data5[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & data5[paste("logFC",contrast.l[2],sep=".")] > logFC &
                               data5[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`101`),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] > logFC & data6[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & data6[paste("logFC",contrast.l[1],sep=".")] > logFC &
                               data6[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`111`),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] > logFC & data7[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & data7[paste("logFC",contrast.l[2],sep=".")] > logFC &
                               data7[paste("adj.P.Val", contrast.l[2],sep=".")] < padj & data7[paste("logFC",contrast.l[1],sep=".")] > logFC &
                               data7[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          }
          
          wb <- createWorkbook() #no li agraden els espais al nom o noms llargs!
          addWorksheet(wb, sheetName = "yellow") 
          writeData(wb,data1.f[!is.na(data1.f[,colnmes[2]]) & 
                                 (data1.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]),
                               cNames], 
                    sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue")
          writeData(wb,data2.f[!is.na(data2.f[,colnmes[2]]) & 
                                 (data2.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]),
                               cNames], 
                    sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data3.f[!is.na(data3.f[,colnmes[2]]) & 
                                 (data3.f[,colnmes[1]] %in% listG3_up[,colnmes[1]]),
                               cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange")
          writeData(wb,data4.f[!is.na(data4.f[,colnmes[2]]) & 
                                 ((data4.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) |
                                    (data4.f[,colnmes[1]] %in% listG2_up[,colnmes[1]])),
                               cNames], 
                    sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data5.f[!is.na(data5.f[,colnmes[2]])  & 
                                 ((data5.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]) |
                                    (data5.f[,colnmes[1]] %in% listG3_up[,colnmes[1]])),
                               cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data6.f[!is.na(data6.f[,colnmes[2]]) & 
                                 ((data6.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) |
                                    (data6.f[,colnmes[1]] %in% listG3_up[,colnmes[1]])),
                               cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "Common brown")
          writeData(wb,data7.f[!is.na(data7.f[,colnmes[2]]) &
                                 ((data7.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG3_up[,colnmes[1]])),cNames], 
                    sheet = "Common brown", startRow = 1, startCol = 1, headerStyle = hs1)
          saveWorkbook(wb,file=file.path(resultsDir,paste("UP.VennGeneLists",filename,"xlsx",sep=".")),overwrite = TRUE)
          
          vennData <- c(length(unique(data6.f$Symbol)),length(unique(data4.f$Symbol)), length(unique(data2.f$Symbol)), length(unique(data3.f$Symbol)),
                        length(unique(data5.f$Symbol)),length(unique(data1.f$Symbol)),
                        length(unique(data7.f$Symbol)))
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("UP.VennDiagram",filename,"png",sep=".")))
          }else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("UP.VennDiagram",filename,"pdf",sep=".")))
          }
          plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), Colors=cols, Title="", shrink=0.5)
          dev.off()
        }else{
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "orange") 
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`100`),
                              cNames], 
                    sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`010`),
                              cNames], 
                    sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`001`),
                              cNames], 
                    sheet = "green", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`110`),
                              cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`011`),
                              cNames], 
                    sheet = "brown", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`101`),
                              cNames], 
                    sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "Common grey")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`111`),
                              cNames], 
                    sheet = "Common grey", startRow = 1, startCol = 1, headerStyle = hs1)
          saveWorkbook(wb,file=file.path(resultsDir,
                                         paste("UP.VennGeneLists",filename,"xlsx",sep=".")), overwrite = TRUE)
          vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
          
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("UP.VennDiagram",filename,"png",sep=".")))
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("UP.VennDiagram",filename,"pdf",sep=".")))
          }
          plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }
      ##### down ######
      cols <- brewer.pal(8,"Pastel2") 
      #Creem l'objecte del Venn
      if (symbols){
        listG1_dn<-listG1_dn[!is.na(listG1_dn[,colnmes[2]]),]
        listG2_dn<-listG2_dn[!is.na(listG2_dn[,colnmes[2]]),]
        listG3_dn<-listG3_dn[!is.na(listG3_dn[,colnmes[2]]),]
        list1<-listG1_dn[,colnmes[2]]
        list2<-listG2_dn[,colnmes[2]]
        list3<-listG3_dn[,colnmes[2]]
      }else{
        list1<-listG1_dn[,colnmes[1]]
        list2<-listG2_dn[,colnmes[1]]
        list3<-listG3_dn[,colnmes[1]]
      }
      list.venn<-list(list1,list2,list3)
      names(list.venn)<-c(listNames[1],listNames[2],listNames[3])
      vtest<-Venn(list.venn)
      
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
        #cNames <- colnames(data4T)
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
      if(symbols){
          
          if(pval!=0){
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`100`),]
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] < -logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,]
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`010`),]
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] < -logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`001`),]
            data3.f <-  data3[data3[paste("logFC",contrast.l[3],sep=".")] < -logFC & data3[paste("P.Value", contrast.l[3],sep=".")] < pval,]
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`110`),]
            data4.f <- data4[data4[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4[paste("P.Value", contrast.l[1],sep=".")] < pval & data4[paste("logFC",contrast.l[2],sep=".")] < -logFC &
                               data4[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`011`),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] < -logFC & data5[paste("P.Value", contrast.l[3],sep=".")] < pval & data5[paste("logFC",contrast.l[2],sep=".")] < -logFC &
                               data5[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`101`),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] < -logFC & data6[paste("P.Value", contrast.l[3],sep=".")] < pval & data6[paste("logFC",contrast.l[1],sep=".")] < -logFC &
                               data6[paste("P.Value", contrast.l[1],sep=".")] < pval,]
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`111`),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] < -logFC & data7[paste("P.Value", contrast.l[3],sep=".")] < pval & data7[paste("logFC",contrast.l[2],sep=".")] < -logFC &
                               data7[paste("P.Value", contrast.l[2],sep=".")] < pval & data7[paste("logFC",contrast.l[1],sep=".")] < -logFC &
                               data7[paste("P.Value", contrast.l[1],sep=".")] < pval,]
          }else{
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`100`),]
            data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] < -logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`010`),]
            data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] < -logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`001`),]
            data3.f <-  data3[data3[paste("logFC",contrast.l[3],sep=".")] < -logFC & data3[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`110`),]
            data4.f <- data4[data4[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4[paste("adj.P.Val", contrast.l[1],sep=".")] < padj & data4[paste("logFC",contrast.l[2],sep=".")] < -logFC &
                               data4[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`011`),]
            data5.f <- data5[data5[paste("logFC",contrast.l[3],sep=".")] < -logFC & data5[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & data5[paste("logFC",contrast.l[2],sep=".")] < -logFC &
                               data5[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`101`),]
            data6.f <- data6[data6[paste("logFC",contrast.l[3],sep=".")] < -logFC & data6[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & data6[paste("logFC",contrast.l[1],sep=".")] < -logFC &
                               data6[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`111`),]
            data7.f <- data7[data7[paste("logFC",contrast.l[3],sep=".")] < -logFC & data7[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & data7[paste("logFC",contrast.l[2],sep=".")] < -logFC &
                               data7[paste("adj.P.Val", contrast.l[2],sep=".")] < padj & data7[paste("logFC",contrast.l[1],sep=".")] < -logFC &
                               data7[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          }
          
          wb <- createWorkbook() #no li agraden els espais al nom o noms llargs!
          addWorksheet(wb, sheetName = "yellow") 
          writeData(wb,data1.f[!is.na(data1.f[,colnmes[2]]) & 
                                 (data1.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]),
                               cNames], 
                    sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue")
          writeData(wb,data2.f[!is.na(data2.f[,colnmes[2]]) & 
                                 (data2.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]),
                               cNames], 
                    sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data3.f[!is.na(data3.f[,colnmes[2]]) & 
                                 (data3.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]]),
                               cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange")
          writeData(wb,data4.f[!is.na(data4.f[,colnmes[2]]) & 
                                 ((data4.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) |
                                    (data4.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]])),
                               cNames], 
                    sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data5.f[!is.na(data5.f[,colnmes[2]])  & 
                                 ((data5.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]) |
                                    (data5.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]])),
                               cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data6.f[!is.na(data6.f[,colnmes[2]]) & 
                                 ((data6.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) |
                                    (data6.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]])),
                               cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "Common brown")
          writeData(wb,data7.f[!is.na(data7.f[,colnmes[2]]) &
                                 ((data7.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG3_dn[,colnmes[1]])),cNames], 
                    sheet = "Common brown", startRow = 1, startCol = 1, headerStyle = hs1)
          saveWorkbook(wb,file=file.path(resultsDir,paste("DN.VennGeneLists",filename,"xlsx",sep=".")),overwrite = TRUE)
          vennData <- c(length(unique(data6.f$Symbol)),length(unique(data4.f$Symbol)), length(unique(data2.f$Symbol)), length(unique(data3.f$Symbol)),
                        length(unique(data5.f$Symbol)),length(unique(data1.f$Symbol)),
                        length(unique(data7.f$Symbol)))
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("DN.VennDiagram",filename,"png",sep=".")))
          }else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("DN.VennDiagram",filename,"pdf",sep=".")))
          }
          plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), Colors=cols, Title="", shrink=0.5)
          dev.off()
        }else{
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "orange") 
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`100`),
                              cNames], 
                    sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`010`),
                              cNames], 
                    sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`001`),
                              cNames], 
                    sheet = "green", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`110`),
                              cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`011`),
                              cNames], 
                    sheet = "brown", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`101`),
                              cNames], 
                    sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "Common grey")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`111`),
                              cNames], 
                    sheet = "Common grey", startRow = 1, startCol = 1, headerStyle = hs1)
          saveWorkbook(wb,file=file.path(resultsDir,
                                         paste("DN.VennGeneLists",filename,"xlsx",sep=".")), overwrite = TRUE)
          vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
          
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("DN.VennDiagram",filename,"png",sep=".")))
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("DN.VennDiagram",filename,"pdf",sep=".")))
          }
          plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }
    }else{
      ##### ALL ######
      cols <- brewer.pal(8,"Pastel2") 
      if(pval!=0) {
        listG1 <- data4T[abs(data4T[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2 <- data4T[abs(data4T[paste("logFC",contrast.l[2],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        listG3 <- data4T[abs(data4T[paste("logFC",contrast.l[3],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[3],sep=".")] < pval,]
      }else{
        listG1 <- data4T[abs(data4T[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2 <- data4T[abs(data4T[paste("logFC",contrast.l[2],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        listG3 <- data4T[abs(data4T[paste("logFC",contrast.l[3],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
        
      }
      
      cols <- brewer.pal(8,"Pastel2") 
      
      #Creem l'objecte del Venn
      if(symbols) {
        listG1<-listG1[!is.na(listG1[,colnmes[2]]),]
        listG2<-listG2[!is.na(listG2[,colnmes[2]]),]
        listG3<-listG3[!is.na(listG3[,colnmes[2]]),]
        list1<-listG1[,colnmes[2]]
        list2<-listG2[,colnmes[2]]
        list3<-listG3[,colnmes[2]]
      }else{
        list1<-listG1[,colnmes[1]]
        list2<-listG2[,colnmes[1]]
        list3<-listG3[,colnmes[1]]
      }
      
      list.venn<-list(list1,list2,list3)
      names(list.venn)<-c(listNames[1],listNames[2],listNames[3])
      vtest<-Venn(list.venn)
      
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
        #cNames <- colnames(data4T)
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                           border = "Bottom", fontColour = "white")
      if(symbols) {
          
          if(pval!=0) {
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`100`),]
            data1.f <- data1[abs(data1[paste("logFC",contrast.l[1],sep=".")]) > logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,]
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`010`),]
            data2.f <- data2[abs(data2[paste("logFC",contrast.l[2],sep=".")]) > logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`001`),]
            data3.f <-  data3[abs(data3[paste("logFC",contrast.l[3],sep=".")]) > logFC & data3[paste("P.Value", contrast.l[3],sep=".")] < pval,]
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`110`),]
            data4.f <- data4[abs(data4[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4[paste("P.Value", contrast.l[1],sep=".")] < pval & abs(data4[paste("logFC",contrast.l[2],sep=".")]) > logFC &
                               data4[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`011`),]
            data5.f <- data5[abs(data5[paste("logFC",contrast.l[3],sep=".")]) > logFC & data5[paste("P.Value", contrast.l[3],sep=".")] < pval & abs(data5[paste("logFC",contrast.l[2],sep=".")]) > logFC &
                               data5[paste("P.Value", contrast.l[2],sep=".")] < pval,]
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`101`),]
            data6.f <- data6[abs(data6[paste("logFC",contrast.l[3],sep=".")]) > logFC & data6[paste("P.Value", contrast.l[3],sep=".")] < pval & abs(data6[paste("logFC",contrast.l[1],sep=".")]) > logFC &
                               data6[paste("P.Value", contrast.l[1],sep=".")] < pval,]
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`111`),]
            data7.f <- data7[abs(data7[paste("logFC",contrast.l[3],sep=".")]) > logFC & data7[paste("P.Value", contrast.l[3],sep=".")] < pval & abs(data7[paste("logFC",contrast.l[2],sep=".")]) > logFC &
                               data7[paste("P.Value", contrast.l[2],sep=".")] < pval & abs(data7[paste("logFC",contrast.l[1],sep=".")]) > logFC &
                               data7[paste("P.Value", contrast.l[1],sep=".")] < pval,]
          }else{
            data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`100`),]
            data1.f <- data1[abs(data1[paste("logFC",contrast.l[1],sep=".")]) > logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
            data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`010`),]
            data2.f <- data2[abs(data2[paste("logFC",contrast.l[2],sep=".")]) > logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`001`),]
            data3.f <-  data3[abs(data3[paste("logFC",contrast.l[3],sep=".")]) > logFC & data3[paste("adj.P.Val", contrast.l[3],sep=".")] < padj,]
            data4 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`110`),]
            data4.f <- data4[abs(data4[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4[paste("adj.P.Val", contrast.l[1],sep=".")] < padj & abs(data4[paste("logFC",contrast.l[2],sep=".")]) > logFC &
                               data4[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data5 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`011`),]
            data5.f <- data5[abs(data5[paste("logFC",contrast.l[3],sep=".")]) > logFC & data5[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & abs(data5[paste("logFC",contrast.l[2],sep=".")]) > logFC &
                               data5[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
            data6 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`101`),]
            data6.f <- data6[abs(data6[paste("logFC",contrast.l[3],sep=".")]) > logFC & data6[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & abs(data6[paste("logFC",contrast.l[1],sep=".")]) > logFC &
                               data6[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
            data7 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`111`),]
            data7.f <- data7[abs(data7[paste("logFC",contrast.l[3],sep=".")]) > logFC & data7[paste("adj.P.Val", contrast.l[3],sep=".")] < padj & abs(data7[paste("logFC",contrast.l[2],sep=".")]) > logFC &
                               data7[paste("adj.P.Val", contrast.l[2],sep=".")] < padj & abs(data7[paste("logFC",contrast.l[1],sep=".")]) > logFC &
                               data7[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          }
          
          wb <- createWorkbook() #no li agraden els espais al nom o noms llargs!
          addWorksheet(wb, sheetName = "yellow") 
          writeData(wb,data1.f[!is.na(data1.f[,colnmes[2]]) & 
                                 (data1.f[,colnmes[1]] %in% listG1[,colnmes[1]]),
                               cNames], 
                    sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue")
          writeData(wb,data2.f[!is.na(data2.f[,colnmes[2]]) & 
                                 (data2.f[,colnmes[1]] %in% listG2[,colnmes[1]]),
                               cNames], 
                    sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data3.f[!is.na(data3.f[,colnmes[2]]) & 
                                 (data3.f[,colnmes[1]] %in% listG3[,colnmes[1]]),
                               cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "orange")
          writeData(wb,data4.f[!is.na(data4.f[,colnmes[2]]) & 
                                 ((data4.f[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data4.f[,colnmes[1]] %in% listG2[,colnmes[1]])),
                               cNames], 
                    sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green1")
          writeData(wb,data5.f[!is.na(data5.f[,colnmes[2]])  & 
                                 ((data5.f[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data5.f[,colnmes[1]] %in% listG3[,colnmes[1]])),
                               cNames], 
                    sheet = "green1", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green2")
          writeData(wb,data6.f[!is.na(data6.f[,colnmes[2]]) & 
                                 ((data6.f[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data6.f[,colnmes[1]] %in% listG3[,colnmes[1]])),
                               cNames], 
                    sheet = "green2", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "Common brown")
          writeData(wb,data7.f[!is.na(data7.f[,colnmes[2]]) &
                                 ((data7.f[,colnmes[1]] %in% listG1[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG2[,colnmes[1]]) |
                                    (data7.f[,colnmes[1]] %in% listG3[,colnmes[1]])),cNames], 
                    sheet = "Common brown", startRow = 1, startCol = 1, headerStyle = hs1)
          saveWorkbook(wb,file=file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
          vennData <- c(length(unique(data6.f$Symbol)),length(unique(data4.f$Symbol)), length(unique(data2.f$Symbol)), length(unique(data3.f$Symbol)),
                        length(unique(data5.f$Symbol)),length(unique(data1.f$Symbol)),
                        length(unique(data7.f$Symbol)))
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
          }else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
          }
          plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), Colors=cols, Title="", shrink=0.5)
          dev.off()
        }else {
          wb <- createWorkbook()
          
          addWorksheet(wb, sheetName = "orange") 
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`100`),
                              cNames], 
                    sheet = "orange", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "blue")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`010`),
                              cNames], 
                    sheet = "blue", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "green")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`001`),
                              cNames], 
                    sheet = "green", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "pink")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`110`),
                              cNames], 
                    sheet = "pink", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "brown")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`011`),
                              cNames], 
                    sheet = "brown", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "yellow")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`101`),
                              cNames], 
                    sheet = "yellow", startRow = 1, startCol = 1, headerStyle = hs1)
          addWorksheet(wb, sheetName = "Common grey")
          writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`111`),
                              cNames], 
                    sheet = "Common grey", startRow = 1, startCol = 1, headerStyle = hs1)
          saveWorkbook(wb,file=file.path(resultsDir,
                                         paste("VennGenes",filename,"xlsx",sep=".")), overwrite = TRUE)                                       
          vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
          
          #PLOT VENN
          if(img.fmt == "png") {
            png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
          } else if (img.fmt == "pdf"){
            pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
          }
          plotVenn3d(vennData, labels=c(listNames[1],listNames[2],listNames[3]), 
                     Colors=cols, Title="", shrink=0.5)
          dev.off()
        }
      }
  }
}
