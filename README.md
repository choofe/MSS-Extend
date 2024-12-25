Goals are as follow:
1. take vehicle information like license plate, VIN, driver name and contacts,... and create a page (sheet) with license plate as the name of the sheet where this info and other information will be kept.
2. keep track of the odometer of each vehicle by updating user inputs which will be made periodically (e.g. weekly)
3. keep track of maintenance items selected from a set of maintainable items during vehicle addition. calculate items mileage and if they are due or near due and provide checklists for those vehicle with due items or ones that user wants. 
4. register services made to items and update item milage based on service time odometer. 
5. register repairs made to vehicles info and costs of the repairs.
6. keep track of vehicle's department in charge and changes in responsible department. 
.....
future extension:
provide the ability to save images of repairshop bills and responsibility change letters.
provide calculation of odometer based on working hour.
----------------------------------------------
Here is a short description on forms and code structures
1. Userform1: main UI which faciliates accessing to parts and functions of app.
2. userform2: A form to collect data like items to track for maintenance and basic vehicle info from the user for adding new vehicle.
3. userform3: Form to edit existing vehicle information like track/untrack items.
4. userform4: Showing a calendar like form which populates a month based on the month that date is in. It works based on Gregorian calendar and converts it to Jalali date (Shamsi Date).
5. userform5: Form to collect data about a service carried out on a specific vehicle and register it to vehicle's page.
6. userform6: Form to register vehicle's department in charge to another department.
7. userform7: Form to register new odometer reading on each vehicle. It add textboxes dynamically for each vehicle and updates the current odometer which is base for planning maintenance and creating checklists.
8. userform8: Form to collect data about a repair that is done and over, and registers it to vehicle's page as history.
9. userform9: Form to select which vehicles to create checklists for. (Currently it has no use because of the changed approach to UI)
10. module1: library for doing main tasks and functions which are repeated often. Also functions and subs needed for date converting are in this module.
11. UserFormEvents: class module for events of dynamically created textebox in runtime. tag property is for indexing options.
12. There might be call to open History.xlsx which is based on an approach to retain histories on a different file for clarity. but there was a risk of losing that file or missing during the copying or moving the main file. New approach is to maintain histories just in the vehicle's page(sheet) which is more consistent.
