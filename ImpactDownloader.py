# Impact Downloader script by Nick Furlo 2018-2019
import configparser
import csv
import glob
import itertools
import os
import shutil
import time
import autohandshake
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import messagebox
from autohandshake import InsightsPage, SurveyPage, EventsPage
from tqdm import tqdm


class ImpactDownloader:
    # loads CSV file into urls dictionary
    def load_csv(self):
        with open(input_file_path, mode='r') as infile:
            urls = {rows[0]: rows[1] for rows in itertools.islice(csv.reader(infile), 1, number_of_rows)}
            return urls

    # Deletes downloaded files from the network locations that are more than 7 days old.
    # Substrings file name, creates datetime from substring, compares datetime, deletes or not.
    def delete_old_network_files(self, cutoff):
        now = time.time()
        print("Checking for files older than " + str(cutoff) + " days.")
        deleted = 0
        for path in network_paths:
            # load all fiels objects into "files"
            files = os.listdir(path)
            for xfile in files:
                full_pth = Path(os.path.join(path, xfile))
                if full_pth.exists():
                    try:
                        created = str(full_pth.name)[-14:-4]
                        created_datetime = datetime.strptime(created, "%Y-%m-%d")
                        today = datetime.now().strftime("%Y-%m-%d")
                        today_datetime = datetime.strptime(today, "%Y-%m-%d")
                        difference = today_datetime - created_datetime
                        # delete file if older than cutoff in days
                        if int(difference.total_seconds() / 86400) > cutoff:
                            os.remove(str(full_pth))
                            deleted += 1
                            print(str(xfile) + " was deleted.")
                    except Exception as e:
                        if "does not match format '%Y-%m-%d'" in str(e):
                            os.remove(str(full_pth))
                            deleted += 1
                        else:
                            print("Error Deleting File: " + str(e))
                else:
                    print("No old files found")
        print(str(deleted) + " old files deleted ")

    # Deletes files from csv download path to make sure old downloads are gone.
    def delete_csv_from_download(self):
        try:
            fileList = os.listdir(download_file_path)
            for fileName in fileList:
                try:
                    os.remove(download_file_path + "/" + fileName)
                except Exception as e:
                    print("Could not delete " + str(fileName) + " because: " + str(e))
            print("CSV files deleted from download directory")
        except Exception as e:
            print("Error Deleting Old Files: " + str(e))

    # Copies files into folder with the name located in column A. These folders are in a
    # network location specified by the user.
    def copy_to_network_drive(self):
        copied = 0
        try:
            print("Starting copy to network location:" + str(network_location))
            os.chdir(download_file_path)
            files = glob.glob("*.csv")
        except:
            print("Could not get files to move to network location.")
        try:
            print("files: " + str(files))
            for file in files:
                # fixes path text
                full_path = os.path.join(network_location, file[:-15])
                full_path = full_path.replace("\\", "/")
                print(full_path)
                network_paths.append(full_path)
                # use shutil to copy file to folder
                shutil.copy2(file, full_path)
                print("Copied " + str(file) + " to " + full_path)
                copied += 1
        except Exception as e:
            print("Could not copy files to network location")
            print("Error code: " + str(e))
        print(str(copied) + " files coppied to " + str(network_location))

    # This function drives most of the script. It iterates through the urls and downlaods the data one at at a time.
    def download_all(self, csv_dict):
        global email, school_url, download_file_path, folder, error_count, event_saved_search_name
        with autohandshake.HandshakeSession(school_url, email, None, 3000, "chromedriver.exe",
                                            os.path.abspath(download_file_path)) as browser:
            csv_dict_reverse = {v: k for k, v in csv_dict.items()}
            for i in tqdm(csv_dict):
                try:
                    # Get folder names
                    folder = csv_dict_reverse[csv_dict[i]]

                    # If the URL is an insights page, download with InsightsPage from autohandshake
                    if 'insights_page' in csv_dict.get(i):
                        # Passes reversed URL dictionary at i for folder name
                        filename = folder + datetime.now().strftime("_%Y-%m-%d") + ".csv"
                        insights_page = InsightsPage(csv_dict[i], browser)

                        # Our appointment page has a custom filter. When using custom filters, "All Results" is not
                        # availbe so instead we use download rows 0-9999
                        if (i == "appointment"):
                            insights_page.download_file(download_dir=download_file_path, file_name=filename,
                                                        file_type=autohandshake.FileType.CSV, limit=9999)
                        else:
                            insights_page.download_file(download_dir=download_file_path, file_name=filename,
                                                        file_type=autohandshake.FileType.CSV, max_wait_time=9999999)

                    # If the URL is a survey page, download with SurveyPage from autohandshake
                    elif 'surveys' in csv_dict.get(i):
                        survey_id = str(csv_dict.get(i))[-5:]
                        print("SID: " + survey_id)
                        survey = SurveyPage(survey_id, browser)
                        survey.download_responses(download_file_path)
                        ImpactDownloader.survey_rename(self)

                    # If the URL is an event page, navigate to the url with browser.get and then use
                    # EventsPage.downlaod_event_data() to download the file. The EventPage requires using a save search
                    # and currently has a bug that requires 1 click from the user. This is why I am using a workaround here.
                    elif 'events' in csv_dict.get(i):
                        events_page = EventsPage(browser)
                        browser.get(csv_dict.get(i))
                        events_page.download_event_data(download_file_path, wait_time=5000)
                        ImpactDownloader.event_rename(self)

                    else:
                        print("\ncould not download" + str(i))
                        error_count = error_count + 1
                except Exception as e:
                    print("\nError: " + str(e))
                    error_count = error_count + 1
            # Output the number of erros resulted in data not being downloaded
            print("\nErrors: " + str(error_count))
            return

    # Rename survey files.
    def survey_rename(self):
        print("survey rename")
        surveyLocation = os.path.join(download_file_path, 'survey_response_download*.csv')
        surveyLocation = os.path.normpath(surveyLocation)
        files = glob.glob(surveyLocation)
        print(files)
        for file in files:
            dst = os.path.join(download_file_path, folder + datetime.now().strftime("_%Y-%m-%d") + ".csv")
            dst = os.path.normpath(dst)
            print("Survey Final Name: " + dst)
            os.rename(file, dst)

    # Rename event files.
    def event_rename(self):
        print("event rename")
        eventLocation = os.path.join(download_file_path, 'event_download*.csv')
        eventLocation = os.path.normpath(eventLocation)
        files = glob.glob(eventLocation)
        print(files)
        for file in files:
            dst = os.path.join(download_file_path, folder + datetime.now().strftime("_%Y-%m-%d") + ".csv")
            dst = os.path.normpath(dst)
            print("Event Final Name: " + dst)

            os.rename(file, dst)


# Create or load config file and assign variables.
def load_config():
    # If there is no config file, make one.
    global school_url, email, password, input_file_path, number_of_rows, download_file_path, network_location, log_enabled, days_until_delete, event_saved_search_name
    config = configparser.ConfigParser()
    my_file = Path("ImpactDownloader.config")
    if not my_file.is_file():
        file = open("ImpactDownloader.config", "w+")
        file.write(
            "[DEFAULT]\nEMAIL = \nSCHOOL_URL= \nINPUT_CSV_FILE_PATH = \nNUMBER_OF_ROWS = \nDOWNLOAD_LOCATION = "
            "\nNETWORK_LOCATION =  \nLOG_TO_FILE = \nDELETE_AFTER_DAYS = \n")
        messagebox.showinfo("Warning",
                            "Config file created. Please add a CSV file path and relaunch the program.")
        sys.exit()
    # Read config file and set global variables
    config.read('ImpactDownloader.config')
    email = config['DEFAULT']['EMAIL']
    school_url = config['DEFAULT']['SCHOOL_URL']
    input_file_path = config['DEFAULT']['INPUT_CSV_FILE_PATH']
    number_of_rows = int(config['DEFAULT']['NUMBER_OF_ROWS'])
    download_file_path = config['DEFAULT']['DOWNLOAD_LOCATION']
    network_location = config['DEFAULT']['NETWORK_LOCATION']
    log_enabled = config['DEFAULT']['LOG_TO_FILE']
    days_until_delete = int(config['DEFAULT']['DELETE_AFTER_DAYS'])
    # Quit if there is no CSV path set in the config file
    if input_file_path is "":
        messagebox.showinfo("Warning",
                            "Please enter an input CSV file path into the config file and relaunch the program.")
        sys.exit()


# Deletes files from csv download path to make sure old downloads are gone.
def delete_csv_from_download():
    try:
        fileList = os.listdir(download_file_path)
        for fileName in fileList:
            try:
                os.remove(download_file_path + "/" + fileName)
            except Exception as e:
                print("Could not delete " + str(fileName) + " because: " + str(e))
        print("CSV files deleted from download directory")
    except Exception as e:
        print("Error Deleting Old Files: " + str(e))


def main(impact_downloader):
    print("Main Method: ")
    global input_file_path, number_of_rows, download_count, missed_urls, log_to_file, download_file_path, days_until_delete
    download_count = 0
    download_file_path = ""
    input_file_path = ""
    number_of_rows = 0
    load_config()
    delete_csv_from_download()
    csv = impact_downloader.load_csv()
    impact_downloader.download_all(csv)
    impact_downloader.copy_to_network_drive()
    impact_downloader.delete_old_network_files(days_until_delete)


error_count = 0
download_count = 0
input_file_path = ""
download_file_path = "-1"
network_location = "-1"
number_of_rows = 0
log_enabled = False
missed_urls = {}
global network_paths
network_paths = []
days_until_delete = 0
main(ImpactDownloader())
