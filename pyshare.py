#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

import csv
import os
import platform
import sys
from optparse import OptionParser

import pymongo
from dateutil import parser
from decouple import config
from shareplum import Office365, Site
from shareplum.site import Version

# Colors to outputs
BGreen = "\033[1;32m"  # Bold Green
BYellow = "\033[1;33m"  # Bold Yellow
BPurple = "\033[1;35m"  # Bold Purple
Yellow = "\033[0;33m"  # Yellow
Blue = "\033[0;34m"  # Blue
Green = "\033[0;32m"  # Green
Red = "\033[0;31m"  # Red

# Background
On_Black = "\033[40m"  # Black Background


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
        self.getData = None
        self.authSpCookie = None
        self.authSpSite = None
        self.authSpList = None

    def __repr__(self):
        return "<" + self.authSpList + ">"

    def auth(self):
        # Authentication
        try:
            print(f"{Blue}Authenticating...")

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

            print(f"{Green}Successfully authenticated!")
        except Exception as e:
            print(f"{Red}SharePoint authentication failed.", e)
            sys.exit(0)

    def get(self):
        # Get items from the Lists
        try:
            print(f"{Blue}Getting items from SharePoint...")

            # Get the list of the site
            self.getData = self.authSpList.GetListItems("All Items")

            print(f"{Green}SharePoint data successfully obtained:")

            for item in self.getData:
                print(f"{Yellow}", item)
        except Exception as e:
            print(f"{Red}Failed getting SharePoint Lists.", e)

    def download(self):
        # Download data as csv from the Lists
        try:
            print(f"{Blue}Downloading csv...")
            path = f"./reports/{self.spList}.csv"

            # Get existing data from SharePoint
            self.get()

            # Download
            with open(f"{path}", "w", encoding="UTF8", newline="") as f:
                writer = csv.DictWriter(f)
                writer.writeheader()
                writer.writerows(self.getData)

            print(f"{Green}Successfully downloaded data from SharePoint!")
            print(f"{Green}Report saved to {path}")
        except Exception as e:
            print(f"{Red}Failed downloading SharePoint Lists.", e)

    def insert(self, path):
        # Insert data from a worksheet file to Microsoft List
        try:
            print(f"{Blue}Reading and inserting data...")

            # Get existing data from SharePoint
            self.get()

            # Insert
            newData = list()
            updateData = list()

            with open(path) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=",")
                fields = next(csv_reader)

                for values in csv_reader:
                    dictionary = dict(zip(fields, values))

                    for data in self.getData:
                        if data["Title"] == dictionary["Title"]:
                            dictionary["ID"] = data["ID"]

                    updateData.append(
                        dictionary
                    ) if "ID" in dictionary else newData.append(dictionary)

            inserted = (
                self.authSpList.UpdateListItems(data=newData, kind="New")
                if len(newData) > 0
                else True
            )
            updated = (
                self.authSpList.UpdateListItems(data=updateData, kind="Update")
                if len(updateData) > 0
                else True
            )

            if inserted and updated:
                print(
                    f"{Green}Successfully inserted {len(newData)} items and updated {len(updateData)} items in the SharePoint!"
                )
        except Exception as e:
            print(f"{Red}Failed inserting data in the SharePoint.", e)

    def create(self, data):
        # Create new items
        try:
            print(f"{Blue}Creating SP items...")
            if len(data) == 0:
                print(f"{Yellow}No data to create!")
                return
            created = self.authSpList.UpdateListItems(data=data, kind="New")
            if created:
                print(f"{Green}Successfully created items!")
        except Exception as e:
            print(f"{Red}SharePoint Lists creation failed.", e)

    def update(self, data):
        # Update the list
        try:
            print(f"{Blue}Updating SP items...")
            if len(data) == 0:
                print(f"{Yellow}No data to update!")
                return
            updated = self.authSpList.UpdateListItems(data=data, kind="Update")
            if updated:
                print(f"{Green}Successfully updated items!")
        except Exception as e:
            print(f"{Red}SharePoint Lists update failed.", e)

    def remove(self, data):
        # Remove items
        try:
            print(f"{Blue}Removing SP items...")
            if len(data) == 0:
                print(f"{Yellow}No data to remove!")
                return
            removed = self.authSpList.UpdateListItems(data=data, kind="Delete")
            if removed:
                print(f"{Green}Successfully removed items!")
        except Exception as e:
            print(f"{Red}SharePoint Lists remove failed.", e)

    def mongoConnect(self):
        # MongoDB connection
        try:
            print(f"{Blue}Connecting to MongoDB...")

            self.mongoClient = pymongo.MongoClient(self.mongoClient)
            self.mongoDatabase = self.mongoClient[self.mongoDatabase]
            self.mongoCollection = self.mongoDatabase[self.mongoCollection]

            print(f"{Yellow}Connected:", self.mongoClient.server_info())
            print(f"{Green}Successfully connected to MongoDB!")
        except Exception as e:
            print(f"{Red}Unable to connect to the MongoDB server.", e)
            sys.exit(0)

    def mongoProcess(self, createData, updateData, connect=True):
        # MongoDB test process
        try:
            # Connect to MongoDB
            if connect:
                self.mongoConnect()

            print(f"{Blue}Running MongoDB test process...")

            if len(createData) > 0:
                print("Creating Mongo data...")
                self.mongoCollection.insert_many(createData)
            else:
                print(f"{Yellow}No data to create!")

            if len(updateData) > 0:
                print("Updating Mongo data...")
                for item in updateData:
                    self.mongoCollection.update_one(
                        {"_id": item["_id"]}, {"$set": item}, upsert=False
                    )
            else:
                print(f"{Yellow}No data to update!")

            items = self.mongoCollection.find({})
            for item in items:
                print(f"{Yellow}", item)

            print(f"{Green}MongoDB test process successfully finished!")
        except Exception as e:
            print(f"{Red}MongoDB test process failed.", e)

    def sync(self):
        # Sync MongoDB with SharePoint Lists
        try:
            print(f"{Blue}Syncing databases...")

            # Connect to MongoDB
            self.mongoConnect()

            # Get existing data from SharePoint
            self.get()
            spData = self.getData

            # Get data from both databases
            mongoData = list(self.mongoCollection.find({}))
            print(f"{Blue}MongoDB data:")
            for item in mongoData:
                print(f"{Yellow}", item)

            addToMongo = list()
            updateToMongo = list()
            addToSp = mongoData.copy()
            updateToSp = list()

            for spItem in spData:
                foundInMongo = False

                for mongoItem in mongoData:
                    # Item in the SP and in the MongoDB
                    if (
                        mongoItem["Title"] == spItem["Title"]
                        and mongoItem["Organization"] == spItem["Organization"]
                    ):
                        foundInMongo = True
                        if mongoItem in addToSp:
                            addToSp.remove(mongoItem)

                        if mongoItem["UpdatedAt"] > spItem["Modificado"]:
                            # Mongo has the newer version
                            item = mongoItem.copy()
                            item["Modificado"] = item["UpdatedAt"]
                            item.pop("UpdatedAt", None)
                            item["ID"] = spItem["ID"]
                            item.pop("_id", None)
                            updateToSp.append(item)
                        elif mongoItem["UpdatedAt"] < spItem["Modificado"]:
                            # SP has the newer version
                            item = spItem.copy()
                            item["_id"] = mongoItem["_id"]
                            item["UpdatedAt"] = item["Modificado"]
                            item.pop("Modificado", None)
                            item.pop("ID", None)
                            updateToMongo.append(item)
                        else:
                            # Same item
                            pass

                # Item only in the SP
                if not foundInMongo:
                    item = spItem.copy()
                    item["UpdatedAt"] = item["Modificado"]
                    item.pop("Modificado", None)
                    item.pop("ID", None)
                    addToMongo.append(item)

            for item in addToSp:
                item["Modificado"] = item["UpdatedAt"]
                item.pop("UpdatedAt", None)
                item.pop("_id", None)

            print(f"{Green} Adding to Mongo: ", addToMongo)
            print(f"{Yellow} Updating to Mongo: ", updateToMongo)
            print(f"{Blue} Adding to SP: ", addToSp)
            print(f"{Red} Updating to SP: ", updateToSp)

            self.mongoProcess(addToMongo, updateToMongo, False)
            self.create(addToSp)
            self.update(updateToSp)

            print(f"{Green}Successfully synced the databases!")
        except Exception as e:
            print(f"{Red}Failed syncing databases.", e)


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
            help="insert data in the SharePoint from a worksheet file",
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

        checkColors()

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
            sharepoint.download()

        # Insert worksheet file data to SharePoint
        if opts.spInsert:
            sharepoint.insert(opts.spInsert)

        # Sync MongoDB and SharePoint
        if opts.spSync:
            sharepoint.sync()
    except Exception as e:
        print(f"{Red}Error:", e)


def showHelp():
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


def checkColors():
    global BGreen
    global BYellow
    global BPurple
    global BCyan
    global Yellow
    global Green
    global Red
    global Blue
    global On_Black

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


if __name__ == "__main__":
    main()
