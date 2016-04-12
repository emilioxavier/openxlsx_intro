Introduction to openxlsx
==================
A quick introduction to the excellent [R](http://www.r-project.org/) 
 package [openxlsx](https://github.com/awalker89/openxlsx) was presented at the
 April 11, 2016 [Lansing Area R User Group](http://www.meetup.com/Lansing-Area-R-Users-Group/) meeting.
 This introduction includes examples demonstrating:  
* Exporting a data.frame  
* Including a stock image  
* Including a ggplot2 image  
* Including a formula  
* Conditional formatting  
* Numerical formatting  

R code with the resulting xlsx file and slides are provided.


***
### Getting openxlsx
The **current and stable** version of [openxlsx](https://cran.r-project.org/web/packages/openxlsx/) 
is available on [CRAN](https://cran.r-project.org) and is installed using the command:
```R
install.packages("openxlsx", dependencies=TRUE)
```

The **development** version of [openxlsx](https://github.com/awalker89/openxlsx) is
available on [GitHub](https://github.com) and is installed using the command:
```R
install.packages(c("Rcpp", "devtools"), dependencies=TRUE)
require(devtools)
install_github("awalker89/openxlsx")
```


