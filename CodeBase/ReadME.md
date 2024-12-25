Here is a short description on forms structures
Userform1: main UI which faciliates accessing to parts and functions of app.
userform2: A form to collect data like items to track for maintenance and basic vehicle info from the user for adding new vehicle.
userform3: Form to edit existing vehicle information like track/untrack items.
userform4: Showing a calendar like form which populates a month based on the month that date is in. It works based on Gregorian calendar and converts it to Jalali date (Shamsi Date).
userform5: Form to collect data about a service carried out on a specific vehicle and register it to vehicle's page.
userform6: Form to register vehicle's department in charge to another department.
userform7: Form to register new odometer reading on each vehicle. It add textboxes dynamically for each vehicle and updates the current odometer which is base for planning maintenance and creating checklists.
userform8: Form to collect data about a repair that is done and over, and registers it to vehicle's page as history.
userform9: Form to select which vehicles to create checklists for. (Currently it has no use because of the changed approach to UI)
module1: library for doing main tasks and functions which are repeated often. Also functions and subs needed for date converting are in this module.
UserFormEvents: class module for events of dynamically created textebox in runtime. tag property
