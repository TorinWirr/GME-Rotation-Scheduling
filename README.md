# GME-Rotation-Scheduling
Using VBA this allows you to draft Resident/Fellow schedules based on templates, trainee requests, program requirements, and available rotations.

## Schedule Templates
The example I'm using is a program with 13 periods and 90 trainees. The schedule template is where you determine which rotation type each trainee is eligible for in each period. You can alter the code to include as many rotation types as you need, for this example I'm just using 2: Inpatient and Subspecialty. Other rotation types might include electives, research, local-away, international, etc.

## Subspecialty Requests/Requirements
Each trainee is allowed to make some scheduling requests. Columns 3-5 in this sheet are where the requests are located. If the program has more than 1 site, specific sites can be chosen here. The remaining columns are the experiences required for program completion/graduation. These are generally not site specific. So, for example, if you have Rheumatology at 3 sites, this could assign them to any of the 3 sites. 

## Subspecialty Inventory
This is where each rotation option is listed and which periods they are available for.

## Inpatient Requests/Requirements
Works same as Subspecialty Requests/Requirements, but with the inpatient rotations

## Inpatient Inventory
Works the same as the Subspecialty Inventory, but with inpatient rotations.

## Basic Flow
Pressing the Inpatient Requests button will cause the program to look at the contents of Period 1 in the first trainees schedule. If it says "Inpatient", then it will look at that trainees Inpatient Requests, then it will look in the Inpatient Inventory for Period 1 to see if that requested rotation is available.

If that rotation is available, it will place that rotation in the schedule template in Period 1 for the trainee, replacing the word "Inpatient". Then, it will take that trainees name and replace the rotation name in the Inpatient Inventory. Last, it will add ": Scheduled" after the rotation name in the Inpatient Requests.

Next it moves to Period 1 for the second trainee, and so on.

## Variation w/Random Sort
The Inpatient button has an additional feature. Each time it finishes looking through a period, it will sort the Inpatient Inventory instead of going through it in the exact same order. This could be applied to each inventory and each button. The reason this feature is added is so you can get different results each time you run it. If the schedule templates/requests/requirements/inventories are always in the same order, you will always get the same results. Adding a random sorting between each period lookup means you can run press all 4 buttons, save it as a new file, go back to the original, press the 4 buttons again, and get a different draft. You can do this numerous times to choose which draft you want to start tweeking manually. A formula in column 1 of the sheet provides the random number to sort by, which is altered after each sort. The sorting itself is in the VBA code attached to the button.

