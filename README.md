# csv2activities #

Application made in winform/c# to connect a third application to SAP b1 instance. 
This application can:

*upload, from a formatted CSV, some activities in CRM->Activities.  
*export today activities in a XML file.
*update, from a formatted CSV, some activities.

## How to format CSV ##

Two files are provided, importAct.csv and updateAct.csv. 
Same structure for both:

*CardCode field ( a code like C23900 )	
*Detail ( string for detail ) 
*Activity ( int, defines the activity primary typology like task, phone call an so. )
*ActivityType (int, defines the secondary typology of the activity, is the ext key for CntctType for OCLT )
*Notes ( string for notes )	
*StartTime ( starting time for the activity in B1 format, like 1030 )	
*EndTime (ending time for the activity in B1 format, like 1730)


