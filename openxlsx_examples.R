library(openxlsx)
library(ggplot2)
library(reshape2)

## cell styles -----------------------------------------------------------------
##----- cell colors
cs.green <- createStyle(textDecoration="bold", fontColour="#006100",
                        bgFill="#c6efce", fgFill="#c6efce")

cs.pink <- createStyle(textDecoration="bold", fontColour="#9c0006",
                       bgFill="#ffc7ce", fgFill="#ffc7ce")

cs.amber <- createStyle(textDecoration="bold", fontColour="darkgoldenrod",
                        bgFill="#ffed98", fgFill="#ffed98")


##----- number of digits
cs.0digits <- createStyle(numFmt="GENERAL")

cs.1digits <- createStyle(numFmt=paste0("0.", paste0(rep(0, 1), collapse="")) )

cs.2digits <- createStyle(numFmt=paste0("0.", paste0(rep(0, 2), collapse="")) )

cs.3digits <- createStyle(numFmt=paste0("0.", paste0(rep(0, 3), collapse="")) )

cs.4digits <- createStyle(numFmt=paste0("0.", paste0(rep(0, 4), collapse="")) )


##----- headers and titles
cs.header <- createStyle(textDecoration="bold",
                         halign="center",
                         border="Bottom")

cs.rows <- createStyle(textDecoration="bold",
                       halign="center",
                       border="Right")

cs.titles.tables <- createStyle(textDecoration="bold",
                                halign="center")


## the iris example ------------------------------------------------------------
##----- calculate the mean and standard deviation for each species
iris.list <- list(iris$Species)

iris.mu <- aggregate(iris, iris.list, mean)
iris.sd <- aggregate(iris, iris.list, sd)

##----- distance based classification
iris.class <- c(rep(1, 50), rep(2, 50), rep(3, 50))
##--- calculate distance
iris.distances <- dist(iris[, 1:4], method="euclidean", diag=FALSE, upper=FALSE)
iris.hclust <- hclust(iris.distances, method="complete")
##--- predict
iris.pred <- cutree(iris.hclust, h=2.25)
##--- evaluate prediction
correct.tf <- iris.class == iris.pred

##----- create results data.frame for the workbook
iris.class <- cbind(iris,
                    iris.class,
                    iris.pred,
                    correct.tf,
                    stringsAsFactors=FALSE)


## the openxlsx example --------------------------------------------------------
results.wb <- createWorkbook()


## summary statistics  ---------------------------------------------------------
addWorksheet(wb=results.wb, sheetName="StatsSummary", gridLines=TRUE)


##----- add data
writeDataTable(wb=results.wb, sheet="StatsSummary",
               x=iris.mu[, -6],
               startCol=1, startRow=3,
               colNames=TRUE, rowNames=FALSE,
               keepNA=FALSE, withFilter=FALSE,
               tableStyle="none",
               headerStyle=NULL)

writeDataTable(wb=results.wb, sheet="StatsSummary",
               x=iris.sd[, -6],
               startCol=1, startRow=9,
               colNames=TRUE, rowNames=FALSE,
               keepNA=FALSE, withFilter=FALSE,
               tableStyle="none",
               headerStyle=NULL)


##----- merge cells for titles
mergeCells(wb=results.wb, sheet="StatsSummary",
           cols=2:5, rows=2)
mergeCells(wb=results.wb, sheet="StatsSummary",
           cols=2:5, rows=8)


##----- add analysis headers
writeData(wb=results.wb, sheet="StatsSummary",
          startCol=2, startRow=2, x="MEAN")
writeData(wb=results.wb, sheet="StatsSummary",
          startCol=2, startRow=8, x="STANDARD DEVIATION")


##----- set column widths
setColWidths(wb=results.wb, sheet="StatsSummary",
             cols=1:5, widths="auto")


##----- number of digits for "Average conservation"
addStyle(wb=results.wb, sheet="StatsSummary", style=cs.2digits,
         cols=2:5, rows=4:6,
         gridExpand=TRUE, stack=TRUE)


##----- number of digits for percentage of clusters with various proportions
addStyle(wb=results.wb, sheet="StatsSummary", style=cs.3digits,
         cols=2:5, rows=10:12,
         gridExpand=TRUE, stack=TRUE)


##----- format the headers
addStyle(wb=results.wb, sheet="StatsSummary", style=cs.header,
         cols=c(2:5), rows=c(2,3,8,9),
         gridExpand=TRUE, stack=TRUE)


## add initial data set --------------------------------------------------------
addWorksheet(wb=results.wb, sheetName="IrisDataSet", gridLines=TRUE)

df.size <- dim(iris.class)
df.row.idc <- 2:(df.size[1] + 1L)

##----- add data
writeDataTable(wb=results.wb, sheet="IrisDataSet",
               x=iris.class,
               startCol=1, startRow=1,
               colNames=TRUE, rowNames=FALSE,
               keepNA=FALSE, withFilter=TRUE,
               tableStyle="none",
               headerStyle=cs.header)

##----- freeze the panes
freezePane(wb=results.wb, sheet="IrisDataSet",
           firstRow=TRUE, firstCol=FALSE)

##----- conditional cell highlighting
conditionalFormatting(wb=results.wb, sheet="IrisDataSet",
                      cols=8, rows=df.row.idc,
                      type="contains", rule="TRUE",
                      style=cs.green)
conditionalFormatting(wb=results.wb, sheet="IrisDataSet",
                      cols=8, rows=df.row.idc,
                      type="contains", rule="FALSE",
                      style=cs.pink)

##----- add a simple formula for a single row of values
the.formula <- c("(A2 + 2*B2) - (C2 + 2*D2)")
class(the.formula) <- c(class(the.formula), "formula")

writeFormula(wb=results.wb, sheet="IrisDataSet",
             x=the.formula,
             startCol=10, startRow=2)


## add a formula over multiple rows --------------------------------------------
addWorksheet(wb=results.wb, sheetName="IrisDataSet_FORMULA", gridLines=TRUE)


##----- construct the formula
part1 <- paste(paste0("A", df.row.idc), paste0("2*B", df.row.idc), sep=" + ")
part2 <- paste(paste0("C", df.row.idc), paste0("2/D", df.row.idc), sep=" + ")
the.formula <- paste(part1, part2, sep=" - ")
class(the.formula) <- c(class(the.formula), "formula")

iris.class.formula <- data.frame(iris.class,
                                 formula=the.formula,
                                 stringsAsFactors=FALSE)

class(iris.class.formula$formula) <- c(class(iris.class.formula$formula), "formula")


##----- freeze the panes
freezePane(wb=results.wb, sheet="IrisDataSet_FORMULA",
           firstRow=TRUE, firstCol=FALSE)


##----- conditional cell highlighting
conditionalFormatting(wb=results.wb, sheet="IrisDataSet_FORMULA",
                      cols=8, rows=df.row.idc,
                      type="contains", rule="TRUE",
                      style=cs.green)
conditionalFormatting(wb=results.wb, sheet="IrisDataSet_FORMULA",
                      cols=8, rows=df.row.idc,
                      type="contains", rule="FALSE",
                      style=cs.pink)


##----- set column widths
setColWidths(wb=results.wb, sheet="IrisDataSet_FORMULA",
             cols=1:9, widths="auto")


##----- format the headers
addStyle(wb=results.wb, sheet="IrisDataSet_FORMULA", style=cs.header,
         cols=c(1:9), rows=1,
         gridExpand=TRUE, stack=TRUE)

writeData(wb=results.wb, sheet="IrisDataSet_FORMULA",
          x=iris.class.formula)

## add an image ----------------------------------------------------------------
##----- add an image
addWorksheet(wb=results.wb, sheetName="theImage", gridLines=TRUE)


insertImage(wb=results.wb, sheet="theImage",
            file="LansingAreaRUserGroup.png",
            startCol=1, startRow=1,
            width=5, height=5,
            units="in", dpi=150)


##----- add a ggplot2 plot
addWorksheet(wb=results.wb, sheetName="ggplot2Plot", gridLines=TRUE)

iris.melt <- melt(iris)
iris.plot <- ggplot(iris.melt, aes(x=variable, y=value, colour=Species))
iris.plot <- iris.plot + geom_line(size=4)
iris.plot <- iris.plot + facet_grid(. ~ Species)

print(iris.plot)
insertPlot(wb=results.wb, sheet="ggplot2Plot",
           fileType="png",
           startCol=1, startRow=1,
           width=15, height=4,
           units="in", dpi=150)


## save the workbook -----------------------------------------------------------
saveWorkbook(results.wb, "openxlsx_examples.xlsx", overwrite=TRUE)


















