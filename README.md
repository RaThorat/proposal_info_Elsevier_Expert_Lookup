# proposal_info_Elsevier_Expert_Lookup
Funding organisations who use Elsevier Expert Lookup for searching reviewers for subsidy applications manually place proposals information per subsidy round in the Expert Lookup. That takes a lot of time and energy. By automatically placing the proposals information in the Expert Lookup, funding organisations can save a lot of time. Elsevier Expert Lookup provides format in which it can take the lumsum proposals information. This code provides missing gap, where the lumsum proposals information from funding organisation is converted into the required format.

## Step : Graphic user interface

This code is a Tkinter GUI application in Python. It has several functions that are called at different points in the application, such as sub_round(), getsuggestedreviewers0(), and country_sheet(). It also has classes defining the layout and behavior of different frames in the GUI, such as StartPage and Page1.

The code starts by importing several modules and setting the current working directory. Then, it defines the functions sub_round(), getsuggestedreviewers0(), and country_sheet().

sub_round() is a function that displays a message box with a message containing user input obtained from a Tkinter Entry widget.

getsuggestedreviewers0() is a function that opens a file selection dialog and reads an Excel file into a Pandas dataframe.

country_sheet() is a function that also opens a file selection dialog and reads an Excel file into a Pandas dataframe.

The code then defines a Tkinter GUI application with several frames, including StartPage and Page1. The StartPage frame has a label and two buttons, one of which allows the user to navigate to the Page1 frame. The Page1 frame has several widgets, including buttons and Entry widgets, which allow the user to interact with the functions defined earlier in the code. The application also includes a "Close" button that allows the user to exit the application.

## Step : Query writing and cleaning
This code is written in SQL and Python. It is used to retrieve specific information from a database and display it in a dataframe. The information that is retrieved includes the file number, main applicant's first and last name, gender, promotion date, email, main organization, organization, postal code, city, title, summary, main discipline, subdiscipline, and a list of words (or "Trefwoorden").

To retrieve this information, the code first defines a variable called "query_code" which includes a SQL query with a placeholder for a value to be entered later. Then, it defines a variable called "Sub_Naam" which is used to store the value for the placeholder in the query. The code then adds a condition to the query to only retrieve information for main applicants and co-applicants.

The code then uses the pyodbc module to connect to the SQL Server and execute the query. The results of the query are then stored in a dataframe called "df_pv". The dataframe is filtered to only include main applicants, and then grouped by certain columns and aggregated to create a new dataframe called "df_pv2". The shape of this dataframe is then printed.

Finally, the code includes a comment indicating that the data in the dataframe can be exported as an Excel file if desired.

## Step : 




