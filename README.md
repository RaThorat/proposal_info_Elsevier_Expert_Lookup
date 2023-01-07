# proposal_info_Elsevier_Expert_Lookup
Funding organisations who use Elsevier Expert Lookup for searching reviewers for subsidy applications manually place proposals information per subsidy round in the Expert Lookup. That takes a lot of time and energy. By automatically placing the proposals information in the Expert Lookup, funding organisations can save a lot of time. Elsevier Expert Lookup provides format in which it can take the lumsum proposals information. This code provides missing gap, where the lumsum proposals information from funding organisation is converted into the required format.

## Graphic user interface

This code is a Tkinter GUI application in Python. It has several functions that are called at different points in the application, such as sub_round(), getsuggestedreviewers0(), and country_sheet(). It also has classes defining the layout and behavior of different frames in the GUI, such as StartPage and Page1.

The code starts by importing several modules and setting the current working directory. Then, it defines the functions sub_round(), getsuggestedreviewers0(), and country_sheet().

sub_round() is a function that displays a message box with a message containing user input obtained from a Tkinter Entry widget.

getsuggestedreviewers0() is a function that opens a file selection dialog and reads an Excel file into a Pandas dataframe.

country_sheet() is a function that also opens a file selection dialog and reads an Excel file into a Pandas dataframe.

The code then defines a Tkinter GUI application with several frames, including StartPage and Page1. The StartPage frame has a label and two buttons, one of which allows the user to navigate to the Page1 frame. The Page1 frame has several widgets, including buttons and Entry widgets, which allow the user to interact with the functions defined earlier in the code. The application also includes a "Close" button that allows the user to exit the application.

