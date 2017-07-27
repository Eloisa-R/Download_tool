# Download_tool
Tool that allows user with an Excel-based freelancer database to programmatically download translation files 
from the platform (by performing http requests).

The tool is used by project managers who have a freelancer and project database based in Excel.
The tool accesses Excel programatically and extracts the necessary information about the project.
Then it makes a http request to the translation platform in order to download the projects
and group them in zip files per assignment.

It checks that all projects are downloaded and alerts the project manager if that's not the case.
It also appends the deadline date before the name of the project to avoid confusion.
