# Resource Locator
----
## Project Overview

During busier times of the year it is necessary to use management resouces in place of traditional employees to help alleviate operational stressors.  The managerial resouces can be scattered across the country at any given point and their ability to help customers is in direct correlation to their geographic distance to those customers for fast and efficient response times and of course asset utilization. 

Using address information for customers, we can convert that data into Latitude and Loingitude as well as home address information of assets and come up with a rolling distance list of assets in relation 
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
 Putting this into an excel file and making some pertinent information hidden was our starting point  
 

 ![res_locator_1](https://github.com/DonChart/Geographical_Distance_Locator_Using_API/assets/168656623/879427eb-638e-4784-bb98-0f5a1ed14ce0)

 Generating a list of employess and their address Lat/Long and indexing it allows us to place a Matrix of hidden data to the right of our user interface to perform some calculations

![res_locator_2](https://github.com/DonChart/Geographical_Distance_Locator_Using_API/assets/168656623/66b84baf-17ce-4390-bcff-bafb45a1278c)  

Those Distance in Miles cells are generated using the Customer Address Lat/Long and comparing it to the Employee Home Address Lat/Long using some advanced math in Excel related to the
Haversine formula - this formula is used to calculate the distance between two points on Earth

----
~~~~
=6371*ACOS(COS(RADIANS(90-$I8))*COS(RADIANS(90-EB$3))+SIN(RADIANS(90-$I8))*SIN(RADIANS(90-EB$3))*COS(RADIANS($J8-EB$4)))/1.609
~~~~
----
### A Quick Breakdown of the formula components:

Radius of the Earth (6371): The Earth's radius is assumed to be 6371 kilometers.

ACOS function: Calculates the arccosine of the value inside it. This is used to compute the central angle between the two points on the Earth's surface.

COS and SIN functions with RADIANS: These functions convert degrees to radians (since trigonometric functions in Excel use radians).

RADIANS(90 - $I8) and RADIANS(90 - EB$3) convert latitude values to radians for point 1 and point 2 respectively.
RADIANS($J8 - EB$4) converts the difference in longitudes between the two points to radians.
Spherical Law of Cosines: The formula inside ACOS calculates the distance using the spherical law of cosines, which is suitable for short distances (e.g., within the same city or region).

Divide by 1.609: This converts the result from kilometers to miles (since 1 kilometer ≈ 0.621371 miles).

Interpretation:
$I8: Latitude of point 1 (in degrees).
EB$3: Latitude of point 2 (in degrees).
$J8: Longitude of point 1 (in degrees).
EB$4: Longitude of point 2 (in degrees).
The formula calculates the distance between the points defined by latitude and longitude coordinates ($I8, $J8) and (EB$3, EB$4) on the Earth's surface, taking into account the curvature of the Earth (using the spherical law of cosines) and converting the result from kilometers to miles.

So, the result of this formula will be the distance between the two points in miles.


Once generated and approprate cells locked down it was copied throughou the matrix
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




## Excel Configuration

Imported data comes into excel completely configured and ready to use as seen here:


![BSC_1](https://github.com/DonChart/Balanced_Scorecard_Estimator_Project/assets/168656623/b4de5a05-db55-488e-808f-b9df4b82d156)


What we do now is some excel trickery using pivot tables and showing/hiding rows using VBA to contain information on one screen for the end user and a button to jump from Weekly/Monthly and   
with the gridlines and row/column headers hidden the transition to tue user is seamless -- to indicate the changes I'll leave the headings on here  

notice - a weekly presentation has rows 21-42 visible and a monthly 60-51 - 

### Weekly Excel Results
----

![BSC_WEEKLY](https://github.com/DonChart/Balanced_Scorecard_Estimator_Project/assets/168656623/f814d978-a00e-47ac-88d2-1dbc0d4f7f0e)

----
### Monthly Excel Results

![BSC_WEEKLY](https://github.com/DonChart/Balanced_Scorecard_Estimator_Project/assets/168656623/8402ce1f-1c69-4d0e-ae87-1eb3ab46613e)

----

### Some VBA

Not a lot of excel VBA required but some creative thinking can make the transition for the end user completely seamless when clicking on a button:
~~~~
Sub HIDE_WEEKLY()

Application.ScreenUpdating = False  ' Hide the Flicker
    Rows("10:46").Select
    Selection.EntireRow.Hidden = True
    Rows("50:88").Select
    Selection.EntireRow.Hidden = False
    Range("A1").Select
End Sub

Sub HIDE_MONTHLY()
Application.ScreenUpdating = False  ' Hide the Flicker
    Rows("10:45").Select
    Selection.EntireRow.Hidden = False
    Rows("45:88").Select
    Selection.EntireRow.Hidden = True
    
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    
    Range("A1").Select
    
End Sub

~~~~

### Formulas
----
Contained within the excel file are some standard formula pulls to get everything in order -- 

#### Vlookups
Grabbing over the data into our final presentation screens  
=IF(ISERROR(VLOOKUP($G60,MONTHLY!$F$2:$LS$1145,K$46,FALSE))," ",VLOOKUP($G60,MONTHLY!$F$2:$LS$1145,K$46,FALSE))  

#### Rankings
Doing some sumproduct ranking to make sure we account for ties and gaps in information  
=SUMPRODUCT((M60<=M$60:M$81)/COUNTIF(M$60:M$81,M$60:M$81))

## Final Thoughts

This particular project was a lot more in depth and complex that I first thought when laying out its requirements with the stakeholders.  It's descriptive seems simple, for every x identity, there will be y number of scorable achievements - add all scores up and rank them over time across the organizational hierarchy.  The back end SQL is the real workhorse of this project because it eliminates a lot of front end Excel calculations.  I attempted to modualize each section so that it would be flexible enough for future changes and uses.  

It would have been far easier to pull raw data into Excel using PowerPivot and DAX calculation information on the fly as the sheet was manipulated by the user but the number crunching across several elements and points in time would have not only blown up the file size when using it, it also would have really affected it's usability and speed.
