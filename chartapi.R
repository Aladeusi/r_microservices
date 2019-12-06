
#Install and load required packages
if (Sys.getenv("JAVA_HOME")!="")
  Sys.setenv(JAVA_HOME="")
library(rJava)

#Install
libraries <-  c("plumber", "knitr")
for(i in libraries) {
  if(i %in% installed.packages() == FALSE) {install.packages(i)}
}

# dependency references

#' @get /get_pie
function(headers, values,title){
  # Create data for the graph.
  
  #convert json parameter to dataframe
  #chart_df=fromJSON(req$postBody)
  hdrs=strsplit(headers, ";")
  vls=strsplit(values, ";")
  
  #converting list to vector and serializing all 
  #the element of the resulting vector to int.
  x <- as.integer(unlist(vls, use.names=FALSE))# c(21, 62,10,53)
  #converting list to vector
  labels <-  unlist(hdrs, use.names=FALSE)#c("London","New York","Singapore","Mumbai")
  
  #converting elements of c to percentage equivalents
  piepercent<- round(100*x/sum(x), 1)
  
  # Give the chart file a name.
  chart_file="C:\\Users\\Training\\Desktop\\piechart.jpg"
  png(file = chart_file)
  
  #Plot the chart.
  pie(x, labels = piepercent, main = title,col = rainbow(length(x)))
  legend("bottom", labels, cex = 0.8,fill = rainbow(length(x)))
  
  # Save the file.
  dev.off()
  
  #reading and convert chart to base 64 image
  library(knitr)
  chart_bs4=image_uri(chart_file)
  #chart_bs4=cat(sprintf("<img src=\"%s\" />\n", chart_bs4))   
  #return bs4 as response
  chart_bs4
  
}

#run the script
#library(plumber)

#pr<-plumber::plumb("C:\\inetpub\\R\\projects\\chartapi.R")
#pr$run(port=8900)




