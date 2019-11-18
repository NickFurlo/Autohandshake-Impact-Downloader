# Impact Downloader script by Nick Furlo 5/21/18
import configparser
import csv
import glob
import itertools
import os
import shutil
import threading
import time
import tkinter
from datetime import datetime
from distutils.util import strtobool
from pathlib import Path
from tkinter import *
from tkinter import messagebox

import keyring as keyring
from autohandshake import HandshakeSession, InsightsPage, SurveyPage, EventsPage
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import autohandshake

print("first")


class ImpactDownloader:

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

    # Copies files into folder with names the same as the name of the file without datetime. These folders are in a
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
            print("Enter try catch")
            print("files: " + str(files))
            for file in files:
                print("enter for loop")
                full_path = os.path.join(network_location, file[:-15])
                full_path = full_path.replace("\\", "/")
                print(full_path)
                network_paths.append(full_path)
                print("Start Shutil")
                shutil.copy2(file, full_path)
                print("Copied File To: " + full_path)
                copied += 1
            print("End of for loop")
        except Exception as e:
            print("Could not copy files to network location")
            print("Error code: " + str(e))
        print(str(copied) + " files coppied to " + str(network_location))

    def download_all(self, csv_dict):
        global email, school_url, download_file_path, folder, error_count, event_saved_search_name
        with autohandshake.HandshakeSession(school_url, email, None, 3000,
                                            None, os.path.abspath(download_file_path)) as browser:
            csv_dict_reverse = {v: k for k, v in csv_dict.items()}
            for i in tqdm(csv_dict):
                try:
                    folder = csv_dict_reverse[csv_dict[i]]
                    if 'insights_page' in csv_dict.get(i):
                        # Passes reversed URL dictionary at i for folder name
                        filename = folder + datetime.now().strftime("_%Y-%m-%d") + ".csv"
                        insights_page = InsightsPage(csv_dict[i], browser)
                        if (i == "appointment"):
                            insights_page.download_file(download_dir=download_file_path, file_name=filename,
                                                        file_type=autohandshake.FileType.CSV, limit=9999)
                        else:
                            insights_page.download_file(download_dir=download_file_path, file_name=filename,
                                                        file_type=autohandshake.FileType.CSV, max_wait_time=9999999)

                    elif 'surveys' in csv_dict.get(i):
                        survey_id = str(csv_dict.get(i))[-5:]
                        print("SID: "+survey_id)
                        survey = SurveyPage(survey_id, browser)
                        survey.download_responses(download_file_path)
                        ImpactDownloader.survey_rename(self)

                    elif 'events' in csv_dict.get(i):
                        events_page = EventsPage(browser)
                        print("\n" + event_saved_search_name)
                        events_page.load_saved_search(event_saved_search_name)
                        #browser.click_element_by_xpath("/html/body/div[2]/div[3]/div/div[2]/div[1]/form/div/div[2]/div/div[3]/div/div[1]/ul/li[1]/div/h3/span")
                        events_page.download_event_data(download_file_path, wait_time=5000)
                        ImpactDownloader.event_rename(self)

                    else:
                        print("\ncould not download" + str(i))
                        error_count = error_count + 1
                except Exception as e:
                    print("\nError: " + str(e))
                    error_count = error_count + 1
            print("\nErrors: " + str(error_count))
            return

    #Rename survey files.
    def survey_rename(self):
        print("survey rename")
        surveyLocation = os.path.join(download_file_path, 'survey_response_download*.csv')
        surveyLocation = os.path.normpath(surveyLocation)
        print("SL: "+ surveyLocation)
        files = glob.glob(surveyLocation)
        print(files)
        for file in files:
            print("Survey File: " + str(file) + "     Folder: " + str(folder))
            dst = os.path.join(download_file_path, folder + datetime.now().strftime("_%Y-%m-%d") + ".csv")
            dst = os.path.normpath(dst)
            os.rename(file, dst)

    #Rename event files.
    def event_rename(self):
        print("event rename")
        eventLocation = os.path.join(download_file_path, 'event_download*.csv')
        eventLocation = os.path.normpath(eventLocation)
        print("EL: "+ eventLocation)
        files = glob.glob(eventLocation)
        print(files)
        for file in files:
            print("Event File: "+str(file)+"     Folder: "+ str(folder))
            dst = os.path.join(download_file_path, folder + datetime.now().strftime("_%Y-%m-%d") + ".csv")
            dst = os.path.normpath(dst)
            os.rename(file,dst)

# Create or load config file and assign variables.
def load_config():
    # If there is no config file, make one.
    global school_url, email, password, input_file_path, number_of_rows, download_file_path, network_location, log_enabled, days_until_delete, event_saved_search_name
    config = configparser.ConfigParser()
    my_file = Path("ImpactDownloader.config")
    if not my_file.is_file():
        file = open("ImpactDownloader.config", "w+")
        file.write(
            "[DEFAULT]\nEMAIL = \nSCHOOL_URL= \nINPUT_CSV_FILE_PATH = \nNUMBER_OF_ROWS = \nEVENT_SAVED_SEARCH_NAME = \nDOWNLOAD_LOCATION = "
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
    event_saved_search_name = str(config['DEFAULT']['EVENT_SAVED_SEARCH_NAME'])
    download_file_path = config['DEFAULT']['DOWNLOAD_LOCATION']
    network_location = config['DEFAULT']['NETWORK_LOCATION']
    log_enabled = config['DEFAULT']['LOG_TO_FILE']
    days_until_delete = config['DEFAULT']['DELETE_AFTER_DAYS']
    # Quit if there is no CSV path set in the config file
    if input_file_path is "":
        messagebox.showinfo("Warning", "Please enter an input CSV file path into the config file and relaunch the program.")
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
    global input_file_path, number_of_rows, download_count, missed_urls, log_to_file, download_file_path
    download_count = 0
    download_file_path = ""
    input_file_path = ""
    number_of_rows = 0
    load_config()
    delete_csv_from_download()
    csv = impact_downloader.load_csv()
    impact_downloader.download_all(csv)
    impact_downloader.copy_to_network_drive()
    # impact_downloader.delete_old_network_files()


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
