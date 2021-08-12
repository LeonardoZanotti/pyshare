#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

import csv

# python modules
import os
import platform
import sys
from optparse import OptionParser

# external modules
import pymongo
from decouple import config
from shareplum import Office365, Site
from shareplum.site import Version


class SharePoint:
    def __init__(self):
        # Set class variables and get environment ones
        self.spLogin = config("SP_LOGIN")
        self.spPassword = config("SP_PASSWORD")
        self.spLink = config("SP_LINK")
        self.spSite = config("SP_SITE")
        self.spList = config("SP_LIST")
        self.mongoClient = config("MONGO_CLIENT")
        self.mongoDatabase = config("MONGO_DATABASE")
        self.mongoCollection = config("MONGO_COLLECTION")
        self.getFields = None
        self.getData = None
        self.authSpCookie = None
        self.authSpSite = None
        self.authSpList = None

    def __repr__(self):
        return "<" + self.authSpList + ">"

    def auth(self):
        # Authentication
        try:
            print("Authenticating...")
            self.authSpCookie = Office365(
                self.spLink, username=self.spLogin, password=self.spPassword
            ).GetCookies()

            # Login in the SharePoint
            self.authSpSite = Site(
                f"{self.spLink}/sites/{self.spSite}",
                version=Version.v365,
                authcookie=self.authSpCookie,
            )

            # Enter the site and get the List
            self.authSpList = self.authSpSite.List(self.spList)
            print("Successfully authenticated!")
        except Exception:
            print("SharePoint authentication failed.", sys.exc_info())

    def get(self):
        # Get items from the Lists
        try:
            print("Getting items from SharePoint...")
            self.getFields = ["ID", "Title"]

            # Get the list of the site
            self.getData = self.authSpList.GetListItems(
                "All Items", fields=self.getFields
            )

            print("Data successfully obtained:")
            for item in self.getData:
                print(item)
        except Exception:
            print("Failed getting SharePoint Lists.", sys.exc_info())

    def download(self):
        # Download data as csv from the Lists
        try:
            print("Downloading csv...")
            # Download
            with open(
                f"./reports/{self.spList}.csv", "w", encoding="UTF8", newline=""
            ) as f:
                writer = csv.DictWriter(f, fieldnames=self.getFields)
                writer.writeheader()
                writer.writerows(self.getData)
            print("Successfully downloaded data from SharePoint!")
        except Exception:
            print("Failed downloading SharePoint Lists.", sys.exc_info())

    def insert(self, path):
        # Insert data from a worksheet file to Microsoft List
        try:
            print("Reading and inserting data...")
            # Insert
            newData = list()
            updateData = list()
            with open(path) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=",")
                fields = next(csv_reader)
                for values in csv_reader:
                    newData.append(dict(zip(fields, values)))
            inserted = self.authSpList.UpdateListItems(data=newData, kind="New")
            if inserted:
                print("Successfully inserted data in the SharePoint!")
        except Exception:
            print("Failed inserting data in the SharePoint.", sys.exc_info())

    def create(self):
        # Create new items
        try:
            # New data to create
            newData = [{"Title": "Bingo"}, {"Title": "Expertise"}]

            print("Creating items...")
            created = self.authSpList.UpdateListItems(data=newData, kind="New")
            if created:
                print("Successfully created items!")
        except Exception:
            print("SharePoint Lists creation failed.", sys.exc_info())

    def update(self):
        # Update the list
        try:
            # Data to update
            updateData = [
                {"ID": "11", "Title": "Belest"},
                {"ID": "12", "Title": "Update 4"},
            ]

            print("Updating items...")
            updated = self.authSpList.UpdateListItems(data=updateData, kind="Update")
            if updated:
                print("Successfully updated items!")
        except Exception:
            print("SharePoint Lists update failed.", sys.exc_info())

    def remove(self):
        # Remove items
        try:
            # Ids to remove
            removeData = ["21"]

            print("Removing items...")
            removed = self.authSpList.UpdateListItems(data=removeData, kind="Delete")
            if removed:
                print("Successfully removed items!")
        except Exception:
            print("SharePoint Lists remove failed.", sys.exc_info())

    def mongo(self):
        # MongoDB
        try:
            print("Connecting to MongoDB...")
            client = pymongo.MongoClient(self.mongoClient)
            print(client.server_info())
            print("Connected!")
            print("Running payload...")
            db = client[self.mongoDatabase]
            collection = db[self.mongoCollection]
            collection.insert_one({"Title": "mongo test"})
            collection.insert_one({"Title": "company two"})
            collection.insert_one({"Title": "company four"})
            collection.update_one(
                {"Title": "company four"}, {"$set": {"Title": "company five"}}
            )
            collection.update_many(
                {"Title": "company five"}, {"$set": {"Title": "company one"}}
            )
            items = collection.find({})
            for item in items:
                print(item)
            collection.delete_many({"Title": "mongo test"})
            print("MongoDB process successfully finished!")
        except Exception:
            print("Unable to connect to the mongo server.", sys.exc_info())

    def sync(self):
        # Sync MongoDB with SharePoint Lists
        try:
            print("Syncing databases...")
            print("Successfully synced the databases!")
        except Exception:
            print("Failed syncing databases.", sys.exc_info())


def main():
    try:
        # Options list
        parser = OptionParser(
            usage="Usage: python3.7 %prog [options]", add_help_option=True
        )
        parser.add_option(
            "-g",
            "--get",
            action="store_true",
            dest="spGet",
            default=False,
            help="list all the items in Microsoft List",
        )
        parser.add_option(
            "-d",
            "--download",
            action="store_true",
            dest="spDownload",
            default=False,
            help="download all the items in Microsoft List as csv worksheet",
        )
        parser.add_option(
            "-i",
            "--insert",
            dest="spInsert",
            metavar="path",
            help="Insert data in the SharePoint from a worksheet file",
        )
        parser.add_option(
            "-c",
            "--create",
            action="store_true",
            dest="spCreate",
            default=False,
            help="create items in Microsoft List",
        )
        parser.add_option(
            "-u",
            "--update",
            action="store_true",
            dest="spUpdate",
            default=False,
            help="update items in Microsoft List",
        )
        parser.add_option(
            "-r",
            "--remove",
            action="store_true",
            dest="spRemove",
            default=False,
            help="remove items of Microsoft List",
        )
        parser.add_option(
            "-m",
            "--mongo",
            action="store_true",
            dest="spMongo",
            default=False,
            help="do some MongoDB test operations",
        )
        parser.add_option(
            "-s",
            "--sync",
            action="store_true",
            dest="spSync",
            default=False,
            help="synchronize the Microsoft List with the MongoDB database",
        )

        opts, args = parser.parse_args()

        # No options passed
        if not any(vars(opts).values()):
            showHelp()
            return

        # Creates new SharePoint instance
        sharepoint = SharePoint()

        # Authentication
        sharepoint.auth()

        # List items
        if opts.spGet:
            sharepoint.get()

        # Download items data as csv
        if opts.spDownload:
            sharepoint.get()
            sharepoint.download()

        # Insert worksheet file data to SharePoint
        if opts.spInsert:
            sharepoint.get()
            sharepoint.insert(opts.spInsert)

        # Create new items
        if opts.spCreate:
            sharepoint.create()

        # Update the list
        if opts.spUpdate:
            sharepoint.update()

        # Remove items
        if opts.spRemove:
            sharepoint.remove()

        # MongoDB
        if opts.spMongo:
            sharepoint.mongo()

        # Sync MongoDB and SharePoint
        if opts.spSync:
            sharepoint.sync()
    except Exception:
        print("Error:", sys.exc_info())


def showHelp():
    # colors
    colors = True  # output colored c:
    machine = sys.platform  # detecting the os
    checkPlatform = platform.platform()  # get current version of os
    if machine.lower().startswith(("os", "win", "darwin", "ios")):
        colors = False  # Mac and Windows shouldn't display colors :c
    if (
        checkPlatform.startswith("Windows-10")
        and int(platform.version().split(".")[2]) >= 10586
    ):
        color = True  # coooolorssss \o/
        os.system("")  # Enables the ANSI -> standard encoding that reads that colors
    if not colors:
        BGreen = BYellow = BPurple = BCyan = Yellow = Green = Red = Blue = On_Black = ""
    else:
        BGreen = "\033[1;32m"  # Bold Green
        BYellow = "\033[1;33m"  # Bold Yellow
        BPurple = "\033[1;35m"  # Bold Purple
        BCyan = "\033[1;36m"  # Bold Cyan
        Yellow = "\033[0;33m"  # Yellow
        Green = "\033[0;32m"  # Green
        Red = "\033[0;31m"  # Red
        Blue = "\033[0;34m"  # Blue

        # Background
        On_Black = "\033[40m"  # Black Background

    print(
        """{BPurple}
    \t                  _                    
    \t  _ __  _   _ ___| |__   __ _ _ __ ___ 
    \t | '_ \| | | / __| '_ \ / _` | '__/ _ \\
    \t | |_) | |_| \__ \ | | | (_| | | |  __/
    \t | .__/ \__, |___/_| |_|\__,_|_|  \___|
    \t |_|    |___/
    \t  
    {BYellow} # Zanotti's SharePoint automation{Blue}
    https://github.com/LeonardoZanotti/pyshare

    To see all the program options use:
    \t
    {BGreen}$ python3.7 pyshare.py -h          
            """.format(
            BPurple=BPurple,
            BGreen=BGreen,
            Blue=Blue,
            BYellow=BYellow,
            On_Black=On_Black,
        )
    )


if __name__ == "__main__":
    main()
