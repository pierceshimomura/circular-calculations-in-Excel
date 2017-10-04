# circular-calculations-in-Excel

Here is a list of the five functions and how to use them:
=circ_r(rng as range)
This function allows you to select a RANGE of cells with 0-360 degree values and it reports back the average vector length (0 to 1). You select a range just like you can select a range with the common Excel function =average(rng). So you can even select a column or disconnected groups of cells. This function reports distinct error values if the avg turns out to be null (no avg for 0 and 180 degrees) or if input values fall outside of the 0-360 range. 
--
=circmean(rng As Range)
This function allows you to select a RANGE of cells with 0-360 degree values and it reports back the average angle in degrees. You select a range just like you can select a range with the common Excel function =average(rng). So you can even select a column or disconnected groups of cells. This function reports distinct error values if the avg turns out to be null (no avg for 0 and 180 degrees) or if input values fall outside of the 0-360 range. 
---
=circ_stdev(rng As Range)
This function allows you to select a RANGE of cells with 0-360 degree values and it reports back the stdev of average angle. You select a range just like you can select a range with the common Excel function =stdev(rng).  This function reports error value if input values fall outside of the 0-360 range. 
---
=absolute_angle(x1, y1, x2, y2) 
This function returns an angle in radians given vector from point (x1, y1) to point (x2, y2). A value of 0 radians is reported for vectors that point straight right. Values increase as you go counterclockwise around a circle until you get to pointing right again. Because of this standard math orientation, you might need to consider how coordinates are in your system. e.g. maybe you want zero degrees to point upward rather than right.
----
=angle(x1, y1, x2, y2, x3, y3, x4, y4) 
This function determines the angle in degrees to turn one vector towards another vector given 4 x,y point coordinates that define the vectors. Intuitively you can think of this angle as the angle that the worm centroid turns while moving. Note that this value is NOT the angle between the two line segments, but is the geometric compliment of that angle. The four points don't need to be nearby each other necessarily. If you want to determine the angle between three points, such as 3 consecutive points in a worm track, make points 2 and 3 the same. The function reports positive values for angles that are counterclockwise and negative values for angles that are clockwise. I put an area on the excel sheet where you can plug in the 8 different input numbers with a graph so that you can see how the function works.
---
 To use these functions with another one of your excel files, you will need to link this visual basic sheet to your own file, or simply copy it and enable macros.
