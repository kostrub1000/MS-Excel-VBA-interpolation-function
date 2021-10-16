# MS-Excel-VBA-interpolation-function
MS Excel/VBA interpolation function
This function enables user of MS Excel interpolate a set of data points (x, y) with a bezier curve drawn through the points of the data set. Instead of looking on a chart, representing the data set and finding values "by hand", one can use this function. It also can find several values of "y" with one "x" if the data set has more than one "y" corresponding to "x". For this user should use array formula in exsel (instead pushing "Enter" one should push "Ctrl+Shift+Enter" on the completion of the formula).
To use this function add the file "spline function.bas" to VBA in your MS Excel file, then type "=spline(x, xRange, yRange)" in a cell and press "Enter" if you want to find one value of "y" or press "Ctrl+Shift+Enter" if you want to find several values of "y", xRange - range of x values, yRange - range of y values that represent the data set that will be interpolated.
The function doesn't extrapolate.
The function doesn't work if xRange, yRange and cell, which user tryes to type the function in, are located in different sheets.
