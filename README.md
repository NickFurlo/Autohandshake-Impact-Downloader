# Autohandshake-Impact-Downloader
Autohandshake Impact Downloader

	This script will download csv files from handshakes insight data, survey results, and event results in bulk. It imports
	the links to the pages through a CSV file and then downloads the data from those links and puts them in a designated
	folder.It can also transfer those files to another location such as a network drive. You can set it to delete files
	older than a user specified amount of time. This is a complete re-write of the HandshakeImpactDownloader. This script
	uses the autohandshake library to handle all webpage interactions. 

Getting Started

	Run the script to create the config file. Edit the config file to include the path of the download link CSV. Run the
	script again and the files will download automatically.

Config File

	The config file will be created upon first run. The following fields will be available.
  
	EMAIL: Email to be entered into handshake for login.
  	SCHOOL_URL: The url for your school's handshake. Ex: https://myschool.joinhandshake.com
	INPUT_CSV_FILE_PATH: Filename or path to csv file containing links
	NUMBER_OF_ROWS: The number of rows including the header. The first row is assumed to be the headers and will not be
	                proccessed
	DOWNLOAD_LOCATION: Path to local download location for CSV's
	NETWORK_LOCATION: Root directory for network location to be copied to. Files will be coppied to subfolders of this
	                  folder with names corresponding with the left column of the CSV. Make sure these folders have been
	                  created prior to running the script.
 	LOG_TO_FILE: True or False to enable logging
  	DELETE_AFTER_DAYS: Looks at the files in the network location folders and deletes all of them older than this amount
  	                   of days

Prerequisites

	Google Chrome
	Python 3.0+
	autohandshake (Python -m pip install autohandshake)
	Selenium (Python -m pip install selenium)
	tqdm (Python -m pip install tqdm)
	python-dateutil (python -m pip install python-dateutil)

CSV Format

   	Row one should consist of headers for the columns. Do not put data in this row.	Column one should contain the name of
	the data. This name is used for the folder and filenames. This will determine the folder that your downloads will be
	transferred too inside of the NETWORK_LOCATION folder. Column two should have the links to insights, event, or survey
	data.

Versioning

	Version 2.0

Authors

	Nick Furlo

License

	The GNU General Public License v3.0
