Calendar Dates
==============

Copyright
---------
Copyright © 2016 TNG Consulting Inc. - http://www.tngconsulting.ca

This file is part of the Calendar Dates application.

Calendar Dates is free software: You can redistribute it and/or modify
it under the terms of the GNU General Public License, version 3,
as published by the Free Software Foundation.

Calendar Dates is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with Calendar Dates. If not, see <http://www.gnu.org/licenses/>.

Authors
-------
Michael Milette, TNG Consulting Inc. - http://www.tngconsulting.ca

Description
-----------
Calendar Dates quickly inserts or removes date numbers in a calendar table.

Requirements and Compatibility
------------------------------
Microsoft Windows (32-bit or 64-bit).

This has been tested with the following applications:
- Microsoft Word
- Microsoft Word Online
- Microsoft OneNote
- Microsoft OneNote Online
- Google Docs
- LibreOffice Writer
- OpenOffice Writer
- WordPerfect
- Lotus Word Pro

Changes
-------
2016-04-10 - Initial version.
For subsequent releases, see CHANGELOG.md.

Installation
------------
No installation is necessary. Just double click on the executable included
in this package:

- CalendarDates32.exe (for 32-bit and 64-bit Windows)
- CalendarDates64.exe (for 64-bit Windows)

Usage & Settings
----------------

CalendarDates doesn't have any configuration settings.

*Adding Dates in a Calendar...*

Use Calendar Dates is a simple 4 step process.

1. Open your document containing an empty calendar table.
   Tables must be exactly 7 columns wide.

2. Launch Calendar Dates. If you don't see your document in the list,
   press F5 or click File > Refresh to update the list of files.

3. Click on the table cell that should contain the 1st of the month.

4. Select the number of days in the month and click the Insert button.
   Tip: You can optionally enter a different number of days than the
   numbers in the drop down pick list. For example, enter 10 will
   send numbers 1 to 10 to the application's table.

Repeat the last 2 steps as often as required. When you are done,
click the Close button.

*Removing Dates from a Calendar...*

If you make a mistake, you can remove dates from the calendar by simply
repeating the above steps but click on the "Delete" button instead of
"Insert".

You can practice on the included CalendarDates-Example.docx file.

Reference
---------

Number of days in each month of a calendar year:

- January = 31 days
- February = 28 days (29 if you can divide the year evenly by 4)
- March = 31 days
- April = 30 days
- May = 31 days
- June = 30 days
- July = 31 days
- August = 31 days
- September = 30 days
- October = 31 days
- November = 30 days
- December = 31 days

Unininstallation
----------------
There is no special uninstallation procedure.

Just delete all the CalendarDates related files.

Security considerations
-----------------------
There are no known security issues at this time.

Motivation for this application
-------------------------------
The development of this application was motivated by the authors need to
create an 18 month calendar but not wanting to type all the dates over
again every year. It is supported by TNG Consulting Inc.

Further information
-------------------
For further information regarding the Calendar Dates application, 
please visit the project page at:

http://github.com/michael-milette/calendar-dates

Language Support
----------------
This application is currently design to work in English and with English
editors. While it may work in other languages, it has not been tested.

This application has not been tested for use with right-to-left (RTL)
languages.

If you want to use this application with a RTL language and it doesn't work
as-is, feel free to prepare a pull request and submit it to the project page
at:

https://github.com/michael-milette/calendar-dates/issues

Limitation
----------
* This tool will only work with the above mentioned applications.
* It does not currently include support other languages.

Known Issues
------------
* The list of documents appearing in the list relies on application titles
which appear in the Windows title bar. As such, titles which do not belong
to a compatible application may occasionally appear in the list.
* For web based word processors, the web browser must not be minimized
and the tab containing the editor must be active in order for the
document to appear in the list of available documents.

Future
------
Have an idea? Wish there were some additional options? 

Share your ideas at:
https://github.com/michael-milette/calendar-dates/issues
