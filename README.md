# xl2csprkrm
my first repository


## Requirements
- **json** (JSON, standard in Python)
- **re** (RegEx, standard in Python)
- **requests** - Go to the [Requests Installation](http://docs.python-requests.org/en/master/user/install/) page and follow the instructions
	- On a mac you could run "sudo pip install requests"
- **openpyxl** - go to the [OpenPyXL documentation](https://openpyxl.readthedocs.org/en/2.3.3/) page and search for "install"
	- On a mac you could run "sudo pip install openpyxl"

## Files
- **test.py** - the code
- **attendees.xlsx** - sample spreadsheet

## Configure script
1. **Set your personal access token**
	- myToken="XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
	- enter your Personal Access token here. You can find it in https://developer.ciscospark.com, login, click your name (top right)
2. **Set your room ID**
	- myRoom="XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
	- get your room ID from the Spark for Developers page, under Rooms / List rooms
3. **Enter the path to your spreadsheet**
	- myworkbook = "/Users/myname/Documents/SparkImport/attendees.xlsx"
4. **Enter the start _column_ that contains the email addresses**
	- example: mysheetStartCol = 3 
	- Enter the column number where the email addresses start:  example: column A=1, B=2, C=3, D=4 ...
5. **Enter the start _row_ that contains the email addresses**
	- example: mysheetStartRow = 4 
	- Enter the row number where the email addresses start
6. **Enable test mode or production**
	- TestOnly = "yes"
	- If "yes": only PRINT results, if "no": add found email addresses to the Spark room.


## Run script
1. Go to the folder containing your test.py file
2. set the "TestOnly" parameter to "yes"
3. run "python test.py"
4. item



## Todo
- Better information on users that could not be added
- Report total number of successful users and failed users



