# SqlQueryApp
### Created a user form to run SQL queries on excel data using access objects, VBA, and SQL.

1. Add a text box that asks for a query name and add a label saying "Query Name" to its left
2. Make two new sheets. i) log sheet with the following columns: (Timestamp,Query Name, Query) ii) Data Dictionary Sheet with (Query Name, Query) as columns
3.  Every time the code runs store the query name in the log sheet with the following columns: (Timestamp,Query Name, Query)
3. Make two list boxes, one for the data dictionary, and one for the activesheets. The listboxes list the query names from both the sheets 
3.a. When a user double clicks a worksheet in the active worksheet listbox then the query from the log sheet is inserted in the Query writing space
3.b. When a user double clicks a worksheet in the data dictionary listbox then the query from the data dictionary sheet is loaded in the Query writing space
4. Other functionalities work as required
