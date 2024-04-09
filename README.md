# Simple VBA code to move or rename files

* Add your current file details and required amendments into the workbook on tab _Input data here_
* Click the button on tab _Initiate_ to start
* The code will execute, and log successes and errors on a new timestamped worksheet called _Run Log_.

Code will not move or rename files if:
* a file with that name and path already exists (to prevent overwriting)
* the file you've specified does not exist

If the new directory you've specified does not exist, the script will create that directory.
