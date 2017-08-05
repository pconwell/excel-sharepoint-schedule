# Worksheets

A script that automatically reads from a schedule (excel file) and enters entries into the worksheet database (sharepoint / MS access)

## Getting Started

1. Install git
2. $ git clone git@github.com:pconwell/worksheets-clean.git
3. ...

### Prerequisites

* Python 3.6
* pypyodbc
* datetime
* openpyxl
* time
* pygit2
* getpass

### Installing

## Running the tests

## Using this program

This program is not too picky, but for it to work correctly there are a few specific things that must be done.

### Patrol schedule

Most importantly, entries on the excel schedule must be entries that are allowed in the worksheet (acces/sharepoint) database. For example, 'A' is for admin leave and is a valid entry in the worksheet database. However, 'C' (just as an example) is not a valid entry in the worksheet database. The following entries are valid:

* A (Admin Leave)
* B (Bereavement)
* F (FMLA)
* H (Holiday)
* JD (Jury Duty)
* M (Military)
* OC (On Call)
* P (Personal Day)
* PNS (Scheduled PTO)
* PNU (Unscheduled PTO)
* SA (Special Assignment)
* SNS (Scheduled Sick)
* SNU (Unscheduled Sick)
* T (Training)
* X (RDO)

Other valid entries are any normal zones that are worked (for example '20' or '34' or '48C'). Any zone that is used on the excel schedule must be in the /config/zones.ini file. Most zones are already entered into this file -- however if in the future a new zone is added, it must be added to the /config/zones.ini file. For example, say zone '15' is added in the future. That zone would need to be added to the /config/zones.ini file. The zones.ini file (or any of the config.ini files for that matter) can be edited in notepad or any other simple text editor.

#### Schedule formatting

The formatting for the excel file should be as follows:

* Each row is an employee
* Each column is a day of the month
* Each worksheet is a month of the year
* The first column (A) can be blank or used for notes. We use our first column to 'color code' days, evenings, and night shift.
* The second column (B) is for categorizing shifts and precincts. For example, B4 might be 'Day Shift', B35 might be 'Evening Shift', etc.
* The third column (C) is the employee's name
* The fourth column (D) is the employee's 3-digit numbers
* The fifth column (E) is the 1st day of the month (the first day that will be entered)
* The first row is the day of week (mon, tue, wed...)
* The second row is the date (1, 2, 3...)
* The third and fourth rows are blank
* The fifth row is the first row with an employee

### Config Files

#### zones.ini

Any zones that are used on the excel schedule and any comments for a given zone can be found in the /config/zones.ini file. The comments can be formatted in one of two ways:

1. `zone = comment`
2. `zone = weekday comment, weekend comment`

For example:

1. `20 = Supervisors`
2. `22 = Central | Traffic, Central`

Notice that during the week, zone 22 has `Traffic`. A comma (" , ") separates the two comments.

Remember, any zone that is used on the patrol schedule must be listed in /config/zones.ini.

#### vehicles.ini

Don't worry about this file for now. It will be fixed up to be more user friendly in a future update. Just for general information, though, the vehicles.ini file holds a list of the vehicles that are available for use.

#### eno.ini

This file includes the 3-digit employee numbers for your assigned employees. Using the template, put each employee's 3-digit ENO on it's own line.

#### config.ini

There are only two things you need to set in this file. Your `shift` and your `precint`. These are the default settings for when you run the program. In other words, this is what the program falls back to if there are no other options. For example, the program knows that zone 32 is part of 'group 2' so the employee would automatically be assigned to that precinct for the day. However, an employee that is on 'A' (admin leave), could be any precinct, so the program falls back to whatever you entered here.

### Worksheet Database

## Contributing

## Authors

* **Patrick Conwell** - *Initial work* - [pconwell](https://github.com/pconwell)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

The sharepoint SOAP API toolbox (sharepoint.py) comes from the https://github.com/jasonrollins/shareplum project.
