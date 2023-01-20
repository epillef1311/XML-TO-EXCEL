# XML-TO-EXCEL
This code reads a folder with xml files in it, read each file searching for the tags specified and writes an excel sheet with headers and cells with the specified tags texts.

If you want to change the tags the code will search for first change the data and __data to what you want(this will be the header of the excel collumn). After that you need to change within the for loop the value of each if ele.tag for the tag you want.

There's a little interface for you to use in case you want.

Note that it may not work exaclty with every xml because it may be diffent the numbers of layers the code goes throught. In this case there's the first layer (root), then a second layer(child) then a third layer(subchild) a fourth layer(subchild2) and a fifith layer(subchild3). Depending of the file you have, there could be more layers or less layers, just add more for loops in case there are more layers or remove the loops in case there is less layers.
