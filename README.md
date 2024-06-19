# Resource Locator
----
## Project Overview

During busier times of the year it is necessary to use management resouces in place of traditional employees to help alleviate operational stressors.  The managerial resouces can be scattered across the country at any given point and their ability to help customers is in direct correlation to their geographic distance to those customers for fast and efficient response times and of course asset utilization. 

Using address information for customers, we can convert that data into Latitude and Longitude as well as home address information of assets and come up with a rolling distance list of assets in relation 
to customers - for this particular project, the stakeholder wanted a 10 person list by ascending distance

Utilizing the Google API - we are also able to plug in any address for other potential needed addresses and a Distance range and pull all emplyees in that range

----
** Disclaimer - This dataset is not active and has some omissions to protect any possible Intellectual Property

## Table of Contents
- [Tools](#Tools)
- [Data and File Preperation](#Data-Perperation-and-Setup)
- [Excel Configuration](#Excel-Configuration)
- [VBA](#Some-Vba)
- [Formulas](#Formulas)
- [Final Thoughts](#Final-Thoughts)


## Data Sources

On Prem T-SQL database, moderately normalized

## Tools
- Excel   | Data Presentation to End User / Primary functional program
- T-Sql   | Data Acqusition
- API     | Google Geocode API

## Data Perperation and Setup

 Having a static list of Account information and address info, using web resources we generated a list of account Lat/Long data - this was our baseline to compare things to
 putting this into an excel file and making some address and lat/lon information hidden from the user was our starting point (shaded fields hidden from user) 
 

 ![res_locator_1](https://github.com/DonChart/Geographical_Distance_Locator_Using_API/assets/168656623/879427eb-638e-4784-bb98-0f5a1ed14ce0)

 Generating a list of employess and their address Lat/Long and indexing it allows us to place a Matrix of hidden data to the right of our user interface cells to perform some calculations

![res_locator_2](https://github.com/DonChart/Geographical_Distance_Locator_Using_API/assets/168656623/66b84baf-17ce-4390-bcff-bafb45a1278c)  

Those Distance in Miles cells are generated using the Customer Address Lat/Long and comparing it to the Employee Home Address Lat/Long using some advanced math in Excel related to the
Haversine formula - this formula is used to calculate the distance between two points on Earth

## Haversine Formula
----
~~~~
=6371*ACOS(COS(RADIANS(90-$I8))*COS(RADIANS(90-EB$3))+SIN(RADIANS(90-$I8))*SIN(RADIANS(90-EB$3))*COS(RADIANS($J8-EB$4)))/1.609
~~~~
----
#### A Quick Breakdown of the formula components:

* Radius of the Earth (6371): The Earth's radius is assumed to be 6371 kilometers.

* ACOS function: Calculates the arccosine of the value inside it. This is used to compute the central angle between the two points on the Earth's surface.

* COS and SIN functions with RADIANS: These functions convert degrees to radians (since trigonometric functions in Excel use radians).

* RADIANS(90 - $I8) and RADIANS(90 - EB$3) convert latitude values to radians for point 1 and point 2 respectively.
  RADIANS($J8 - EB$4) converts the difference in longitudes between the two points to radians.
  Spherical Law of Cosines: The formula inside ACOS calculates the distance using the spherical law of cosines, which is suitable for short distances (e.g., within the same city or region).

* Divide by 1.609: This converts the result from kilometers to miles (since 1 kilometer ≈ 0.621371 miles).

Interpretation:
$I8: Latitude of point 1 (in degrees).
EB$3: Latitude of point 2 (in degrees).
$J8: Longitude of point 1 (in degrees).
EB$4: Longitude of point 2 (in degrees).
The formula calculates the distance between the points defined by latitude and longitude coordinates ($I8, $J8) and (EB$3, EB$4) on the Earth's surface, taking into account the curvature of the Earth (using the spherical law of cosines) and converting the result from kilometers to miles.  

The result of this formula will be the distance between the two points in miles.
Once the formula was tested and generated it was copied throughout the matrix

----

Now that we have a matrix of all employees and their realtive distance to our source address, it a straightforward process to generate who is the closest using some Excel Formula's  

The Distance in Miles for our closest employee (Position 1) @ 17.10 Miles is derived by finding the smallest value across our matrix for that row
````
=SMALL($EB8:$AYL8,1)
````

![res_locator_3](https://github.com/DonChart/Geographical_Distance_Locator_Using_API/assets/168656623/2525f41d-bd4d-4526-b809-0c34db1ca5b7)

Now, in a pair of hidden columns we use some more formulas to Match that smallest distance from our Matrix to the Employee List
To find our Column position in the Matrix
~~~~
=MATCH($L8,$EB8:$AYL8,0)
~~~~
And to pull our index potion in the Matrix
~~~~
=INDEX($EB$2:$AYL$2,$M8)
~~~~
We then use that index position across our employee table to find the Employee Name etc...
~~~~
=VLOOKUP($N8,EMPLOYEE_INFO!$A$3:$P$1411,8,FALSE)
~~~~

----

After position 1 is complete, we duplicate things to position two, however this time we use an array formula to find the next mileage distance in our matrix
~~~~
=MIN(IF($EB8:$AYL8>L8,$EB8:$AYL8))
~~~~
Git Hub does't show curly brackets for the array, but this is in fact an array - it references our Matrix to make sure that the returned value is larger than our cell L8 (17.10 miles)
and returns the next employee at @25.09 miles  

We carry this process across 10 levels to give our stakeholder a complete picture


![res_locator_4](https://github.com/DonChart/Geographical_Distance_Locator_Using_API/assets/168656623/e4f15890-e141-4dcf-b9bb-0d8eda4e2188)





## Final Thoughts

This particular project was a lot more in depth and complex that I first thought when laying out its requirements with the stakeholders.  It's descriptive seems simple, for every x identity, there will be y number of scorable achievements - add all scores up and rank them over time across the organizational hierarchy.  The back end SQL is the real workhorse of this project because it eliminates a lot of front end Excel calculations.  I attempted to modualize each section so that it would be flexible enough for future changes and uses.  

It would have been far easier to pull raw data into Excel using PowerPivot and DAX calculation information on the fly as the sheet was manipulated by the user but the number crunching across several elements and points in time would have not only blown up the file size when using it, it also would have really affected it's usability and speed.
