
Project Name: IMEI and Serial Number Search Script

Description:
This Python-based automated search script retrieves information by searching for IMEI and Serial numbers in a user-specified Excel file on the web. The project utilizes the selenium library to automate interactions with a web page, allowing it to input data and retrieve relevant information. Additionally, it employs the openpyxl library to read and update data in an Excel file.

Usage:

Initializes a web browser using Selenium.
Navigates to the specified website (https://snlookup.com/samsung).
Enters the IMEI or Serial number from a specific cell in the Excel file on the website.
The retrieved information from the website is recorded in the corresponding row in the same Excel file.
Notes:

The project reads IMEI and Serial numbers by scanning a specific column in the Excel file.
Different XPaths are used on the website based on whether the input is an IMEI or Serial number.
In case of an error, the application provides a notification and stops execution.
The script is designed to use the Chrome browser; however, it can be modified for other browsers.
Requirements:

Python 3.x
selenium library
openpyxl library
Chrome browser and chromedriver (or an appropriate driver for the chosen browser)
