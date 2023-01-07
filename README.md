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

## Step : Edit information in the format of example from Expert Lookup
This code is part of a larger project that is intended to edit information from an Excel file called 'Proposal Upload Example.xlsx' and create two new DataFrames: 'Proposal Details' and 'Proposal Applicants'. The 'Proposal Details' DataFrame contains information about grant proposals, including the grant number, council name, proposal title, abstract, specific aims, and keywords. The 'Proposal Applicants' DataFrame contains information about the applicants for each grant proposal, including their name, email, affiliation, country, and role in the proposal.

The code begins by creating the 'Proposal Details' DataFrame, and then assigns values to each of its columns. The 'Grant No.' and 'No' columns are taken from the 'dossiernummer' column of the 'df_pv2' DataFrame, and the 'Council' column is given a fixed value of 'Name of the funder'. The 'Proposal Title', 'Abstract', and 'Keywords' columns are taken directly from the 'df_pv2' DataFrame.

The 'Proposal Applicants' DataFrame is then created and its columns are given values. The 'Grant No.' and 'Proposal Number' columns are taken from the 'dossiernummer' column of the 'df_pv2' DataFrame, and the 'Last Name' and 'First Name' columns are taken from the 'achternaam' and 'voornaam' columns of 'df_pv2', respectively. The 'Email', 'Affiliation', and 'Country' columns are also taken from 'df_pv2', and the 'Role' column is given a fixed value of 'Principal Investigator'. The 'Is Principal Investigator?' column is given a fixed value of 'Yes'.

The code then creates a new DataFrame called 'df_colab_1' by grouping and aggregating the 'df_colab_0' DataFrame and renaming its 'woord' column to 'Trefwoorden'. Finally, the 'df_colab' DataFrame is created and its columns are given values taken from the 'df_colab_1' DataFrame. The 'Role' column is given a fixed value of 'Co-applicant', and the 'Is Principal Investigator?' column is given a fixed value of 'No'.

## Step: Bundeling all information in a single file
This code is used to create an Excel workbook with four worksheets. The first worksheet, "Proposal Details," is created using the data from the Pandas Dataframe "df_PD." The second worksheet, "Proposal Applicants," is created using the data from the Pandas Dataframe "df_PA33." The third worksheet, "Suggested Reviewers," is created using the data from the Pandas Dataframe "df_SR." The fourth worksheet, "Country Sheet," is created using the data from the Pandas Dataframe "df_CS." The data for each worksheet is added to the worksheet using the openpyxl.utils.dataframe.dataframe_to_rows() function. The worksheets are then saved in an Excel workbook with the file name "Proposal info to feed in EL.xlsx."
