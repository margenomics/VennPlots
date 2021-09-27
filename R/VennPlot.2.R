VennPlot.2 <- function(contrast.l , listNames, filename, data4T= NULL, symbols=TRUE, 
                             colnmes= c("AffyID", "Symbol"),listSymbols=F,listA=NULL, listB=NULL, 
                             img.fmt= "pdf", pval=0, padj=0.05, logFC=1, up_down=FALSE){
  if (listSymbols){
    listG1<-listA
    listG2<-listB
    
    cols <- brewer.pal(8,"Pastel2") 
    
    #Ens assegurem de que no hi ha NA
    listG1 <- listG1[!is.na(listG1)]
    listG2 <- listG2[!is.na(listG2)]
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
    
    if(!is.null(data4T)){
      #EXCEL
      #cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      cNames <- colnames(data4T)
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      
      wb <- createWorkbook()
      addWorksheet(wb, sheetName = listNames[1])
      writeData(wb,data4T[(!is.na(data4T[,"Symbol"])) & 
                            (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`10`)), cNames], 
                sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = listNames[2])
      writeData(wb,data4T[(!is.na(data4T[,"Symbol"])) & 
                            (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`01`)), cNames], 
                sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Common")
      writeData(wb,data4T[(!is.na(data4T[,"Symbol"])) & 
                            (data4T[,"Symbol"] %in% unlist(vtest@IntersectionSets$`11`)),cNames],
                sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
      saveWorkbook(wb,file = file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
      
    }else{
      #EXCEL
      genes <- data.frame("Genes"=unlist(list.venn))
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      
      wb <- createWorkbook()
      addWorksheet(wb, sheetName = listNames[1])
      writeData(wb,genes[genes$Genes %in% unlist(vtest@IntersectionSets$`10`),], 
                sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = listNames[2])
      writeData(wb,genes[genes$Genes %in% unlist(vtest@IntersectionSets$`01`),], 
                sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
      addWorksheet(wb, sheetName = "Common")
      writeData(wb,unique(genes[genes$Genes %in% unlist(vtest@IntersectionSets$`11`),]),
                sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
      saveWorkbook(wb,file = file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
      
    }
  }else{
    if (up_down){ ###################### UP & DOWN SEPARATED ############################
      if (pval!=0){
        listG1_up <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2_up <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] > logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
        
        
        listG1_dn <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2_dn <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] < -logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
      }else{
        listG1_up <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2_up <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] > logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        
        
        listG1_dn <- data4T[data4T[paste("logFC",contrast.l[1],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2_dn <- data4T[data4T[paste("logFC",contrast.l[2],sep=".")] < -logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
      }
      
      ##### ##### ##### ##### ##### ##### up ###### ##### ##### ##### ##### #####
      cols <- brewer.pal(8,"Pastel2") 
      #Creem l'objecte del Venn
      if (symbols){
        listG1_up<-listG1_up[!is.na(listG1_up[,colnmes[2]]),]
        listG2_up<-listG2_up[!is.na(listG2_up[,colnmes[2]]),]
        list1<-listG1_up[,colnmes[2]]
        list2<-listG2_up[,colnmes[2]]
      }else{
        list1<-listG1_up[,colnmes[1]]
        list2<-listG2_up[,colnmes[1]]
      }
      list.venn<-list(list1,list2)
      names(list.venn)<-c(listNames[1],listNames[2])
      vtest<-Venn(list.venn)
      
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      #cNames <- colnames(data4T)
      
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      if (symbols) {
        if (pval!=0){
          data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10`),]
          data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] > logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,]
          data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01`),]
          data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] > logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,]
          data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11`),]
          data3.f <- data3[data3[paste("logFC",contrast.l[1],sep=".")] > logFC & data3[paste("P.Value", contrast.l[1],sep=".")] < pval & data3[paste("logFC",contrast.l[2],sep=".")] > logFC &
                             data3[paste("P.Value", contrast.l[2],sep=".")] < pval,]
          
        }else{
          data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10`),]
          data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] > logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01`),]
          data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] > logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
          data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11`),]
          data3.f <- data3[data3[paste("logFC",contrast.l[1],sep=".")] > logFC & data3[paste("adj.P.Val", contrast.l[1],sep=".")] < padj & data3[paste("logFC",contrast.l[2],sep=".")] > logFC &
                             data3[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        }
        wb <- createWorkbook()
        addWorksheet(wb, sheetName = listNames[1])
        writeData(wb,data1.f[(!is.na(data1.f[,colnmes[2]])) &
                               (data1.f[,colnmes[1]] %in% listG1_up[,colnmes[1]]), cNames], 
                  sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = listNames[2])
        writeData(wb,data2.f[(!is.na(data2.f[,colnmes[2]])) & 
                               (data2.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]), cNames], 
                  sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common")
        writeData(wb,data3.f[(!is.na(data3.f[,colnmes[2]])) & 
                               (data3.f[,colnmes[1]] %in% listG1_up[,colnmes[1]] |  #aquí corregit 20/9/16
                                  data3.f[,colnmes[1]] %in% listG2_up[,colnmes[1]]),cNames],
                  sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
        saveWorkbook(wb,file=file.path(resultsDir,paste("UP.VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
        #els noms estan al reves al plot labels
        vennData <- c(length(unique(data1.f$Symbol)), length(unique(data2.f$Symbol)), length(unique(data3.f$Symbol)))
        if(img.fmt == "png") {
          png(file.path(resultsDir,paste("UP.VennDiagram",filename,"png",sep=".")))
        } else if (img.fmt == "pdf"){
          pdf(file.path(resultsDir,paste("UP.VennDiagram",filename,"pdf",sep=".")))
        }
        plotVenn2d(vennData, labels=c(listNames[1],listNames[2]), Colors=cols, Title="", shrink=1)
        dev.off()
      }else {
        wb <- createWorkbook()
        addWorksheet(wb, sheetName = listNames[1])
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10`),
                            cNames], 
                  sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = listNames[2])
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01`),
                            cNames], 
                  sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common")
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11`),
                            cNames],
                  sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
        saveWorkbook(wb,file = file.path(resultsDir,paste("UP.VennGeneLists",filename,"xlsx",sep=".")),overwrite = TRUE)
        #els noms estan al reves al plot labels
        vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
        if(img.fmt == "png") {
          png(file.path(resultsDir,paste("UP.VennDiagram",filename,"png",sep=".")))
        } else if (img.fmt == "pdf"){
          pdf(file.path(resultsDir,paste("UP.VennDiagram",filename,"pdf",sep=".")))
        }
        plotVenn2d(vennData, labels=c(listNames[2],listNames[1]), Colors=cols, Title="", shrink=1)
        dev.off()
      }
      ##### ##### ##### ##### ##### ##### down  ##### ##### ##### ##### #####
      cols <- brewer.pal(8,"Pastel2") 
      #Creem l'objecte del Venn
      if (symbols){
        listG1_dn<-listG1_dn[!is.na(listG1_dn[,colnmes[2]]),]
        listG2_dn<-listG2_dn[!is.na(listG2_dn[,colnmes[2]]),]
        list1<-listG1_dn[,colnmes[2]]
        list2<-listG2_dn[,colnmes[2]]
      }else{
        list1<-listG1_dn[,colnmes[1]]
        list2<-listG2_dn[,colnmes[1]]
      }
      list.venn<-list(list1,list2)
      names(list.venn)<-c(listNames[1],listNames[2])
      vtest<-Venn(list.venn)
      #EXCEL
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      #cNames <- colnames(data4T)
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      if (symbols) {
        
        if (pval!=0) {
          data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10`),]
          data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] < -logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,]
          data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01`),]
          data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] < -logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,]
          data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11`),]
          data3.f <- data3[data3[paste("logFC",contrast.l[1],sep=".")] < -logFC & data3[paste("P.Value", contrast.l[1],sep=".")] < pval &
                             data3[paste("logFC",contrast.l[2],sep=".")] < -logFC & data3[paste("P.Value", contrast.l[2],sep=".")] < pval,]
          
        }else{
          data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10`),]
          data1.f <- data1[data1[paste("logFC",contrast.l[1],sep=".")] < -logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01`),]
          data2.f <- data2[data2[paste("logFC",contrast.l[2],sep=".")] < -logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
          data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11`),]
          data3.f <- data3[data3[paste("logFC",contrast.l[1],sep=".")] < -logFC & data3[paste("adj.P.Val", contrast.l[1],sep=".")] < padj &
                             data3[paste("logFC",contrast.l[2],sep=".")] < -logFC & data3[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        }
        
        wb <- createWorkbook()
        addWorksheet(wb, sheetName = listNames[1])
        writeData(wb,data1.f[(!is.na(data1.f[,colnmes[2]])) &
                               (data1.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]]), cNames], 
                  sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = listNames[2])
        writeData(wb,data2.f[(!is.na(data2.f[,colnmes[2]])) &
                               (data2.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]), cNames], 
                  sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common")
        writeData(wb,data3.f[(!is.na(data3.f[,colnmes[2]])) & 
                               (data3.f[,colnmes[1]] %in% listG1_dn[,colnmes[1]] |  #aquí corregit 20/9/16
                                  data3.f[,colnmes[1]] %in% listG2_dn[,colnmes[1]]),cNames],
                  sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
        saveWorkbook(wb,file=file.path(resultsDir,paste("DN.VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
        #els noms estan al reves al plot labels
        vennData <- c(length(unique(data1.f$Symbol)), length(unique(data2.f$Symbol)), length(unique(data3.f$Symbol)))
        if(img.fmt == "png") {
          png(file.path(resultsDir,paste("DN.VennDiagram",filename,"png",sep=".")))
        } else if (img.fmt == "pdf"){
          pdf(file.path(resultsDir,paste("DN.VennDiagram",filename,"pdf",sep=".")))
        }
        plotVenn2d(vennData, labels=c(listNames[1],listNames[2]), Colors=cols, Title="", shrink=1)
        dev.off()
        
      }else {
        wb <- createWorkbook()
        addWorksheet(wb, sheetName = listNames[1])
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10`),
                            cNames], 
                  sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = listNames[2])
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01`),
                            cNames], 
                  sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common")
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11`),
                            cNames],
                  sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
        saveWorkbook(wb,file = file.path(resultsDir,paste("DN.VennGeneLists",filename,"xlsx",sep=".")),overwrite = TRUE)
        #els noms estan al reves al plot labels
        vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
        if(img.fmt == "png") {
          png(file.path(resultsDir,paste("DN.VennDiagram",filename,"png",sep=".")))
        } else if (img.fmt == "pdf"){
          pdf(file.path(resultsDir,paste("DN.VennDiagram",filename,"pdf",sep=".")))
        }
        plotVenn2d(vennData, labels=c(listNames[2],listNames[1]), Colors=cols, Title="", shrink=1)
        dev.off()
        
      }
    }else{
      if (pval!=0){
        listG1 <- data4T[abs(data4T[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[1],sep=".")] < pval,]
        listG2 <- data4T[abs(data4T[paste("logFC",contrast.l[2],sep=".")]) > logFC & data4T[paste("P.Value", contrast.l[2],sep=".")] < pval,]
      }else{
        listG1 <- data4T[abs(data4T[paste("logFC",contrast.l[1],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
        listG2 <- data4T[abs(data4T[paste("logFC",contrast.l[2],sep=".")]) > logFC & data4T[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
      }
      
      cols <- brewer.pal(8,"Pastel2") 
      
      #Creem l'objecte del Venn
      if (symbols){
        listG1<-listG1[!is.na(listG1[,colnmes[2]]),]
        listG2<-listG2[!is.na(listG2[,colnmes[2]]),]
        list1<-listG1[,colnmes[2]]
        list2<-listG2[,colnmes[2]]
      }else{
        list1<-listG1[,colnmes[1]]
        list2<-listG2[,colnmes[1]]
      }
      
      list.venn<-list(list1,list2)
      names(list.venn)<-c(listNames[1],listNames[2])
      vtest<-Venn(list.venn)
      
      #EXCEL
      cNames <- colnames(data4T)[-grep("scaled$", colnames(data4T))]
      #cNames <- colnames(data4T)
      hs1 <- createStyle(fgFill = "#737373", halign = "CENTER", textDecoration = "Bold",
                         border = "Bottom", fontColour = "white")
      if (symbols){
        
        if (pval!=0){
          data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10`),]
          data1.f <- data1[abs(data1[paste("logFC",contrast.l[1],sep=".")]) > logFC & data1[paste("P.Value", contrast.l[1],sep=".")] < pval,]
          data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01`),]
          data2.f <- data2[abs(data2[paste("logFC",contrast.l[2],sep=".")]) > logFC & data2[paste("P.Value", contrast.l[2],sep=".")] < pval,]
          data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11`),]
          data3.f <- data3[abs(data3[paste("logFC",contrast.l[1],sep=".")]) > logFC & data3[paste("P.Value", contrast.l[1],sep=".")] < pval &
                             abs(data3[paste("logFC",contrast.l[2],sep=".")]) > logFC & data3[paste("P.Value", contrast.l[2],sep=".")] < pval,]
          
        }else{
          data1 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`10`),]
          data1.f <- data1[abs(data1[paste("logFC",contrast.l[1],sep=".")]) > logFC & data1[paste("adj.P.Val", contrast.l[1],sep=".")] < padj,]
          data2 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`01`),]
          data2.f <- data2[abs(data2[paste("logFC",contrast.l[2],sep=".")]) > logFC & data2[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
          data3 <- data4T[data4T[,colnmes[2]] %in% unlist(vtest@IntersectionSets$`11`),]
          data3.f <- data3[abs(data3[paste("logFC",contrast.l[1],sep=".")]) > logFC & data3[paste("adj.P.Val", contrast.l[1],sep=".")] < padj & 
                             abs(data3[paste("logFC",contrast.l[2],sep=".")]) > logFC & data3[paste("adj.P.Val", contrast.l[2],sep=".")] < padj,]
        }
        
        wb <- createWorkbook()
        addWorksheet(wb, sheetName = listNames[1])
        writeData(wb,data1.f[(!is.na(data1.f[,colnmes[2]])) &
                               (data1.f[,colnmes[1]] %in% listG1[,colnmes[1]]), cNames], 
                  sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = listNames[2])
        writeData(wb,data2.f[(!is.na(data2.f[,colnmes[2]])) & 
                               (data2.f[,colnmes[1]] %in% listG2[,colnmes[1]]), cNames], 
                  sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common")
        writeData(wb,data3.f[(!is.na(data3.f[,colnmes[2]])) & 
                               (data3.f[,colnmes[1]] %in% listG1[,colnmes[1]] |  #aquí corregit 20/9/16
                                  data3.f[,colnmes[1]] %in% listG2[,colnmes[1]]),cNames],
                  sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
        saveWorkbook(wb,file=file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
        #els noms estan al reves al plot labels
        vennData <- c(length(unique(data1.f$Symbol)), length(unique(data2.f$Symbol)), length(unique(data3.f$Symbol)))
        if(img.fmt == "png") {
          png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
        } else if (img.fmt == "pdf"){
          pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
        }
        plotVenn2d(vennData, labels=c(listNames[1],listNames[2]), Colors=cols, Title="", shrink=1)
        dev.off()
      }else {
        wb <- createWorkbook()
        addWorksheet(wb, sheetName = listNames[1])
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`10`),
                            cNames], 
                  sheet = listNames[1], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = listNames[2])
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`01`),
                            cNames], 
                  sheet = listNames[2], startRow = 1, startCol = 1, headerStyle = hs1)
        addWorksheet(wb, sheetName = "Common")
        writeData(wb,data4T[data4T[,colnmes[1]] %in% unlist(vtest@IntersectionSets$`11`),
                            cNames],
                  sheet = "Common", startRow = 1, startCol = 1, headerStyle = hs1)
        saveWorkbook(wb,file = file.path(resultsDir,paste("VennGenes",filename,"xlsx",sep=".")),overwrite = TRUE)
        #els noms estan al reves al plot labels
        vennData <- sapply(vtest@IntersectionSets,function(x) length(unlist(x)))
        if(img.fmt == "png") {
          png(file.path(resultsDir,paste("VennDiagram",filename,"png",sep=".")))
        } else if (img.fmt == "pdf"){
          pdf(file.path(resultsDir,paste("VennDiagram",filename,"pdf",sep=".")))
        }
        plotVenn2d(vennData, labels=c(listNames[2],listNames[1]), Colors=cols, Title="", shrink=1)
        dev.off()
      }
    }
  }
}
