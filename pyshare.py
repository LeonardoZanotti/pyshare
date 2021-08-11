#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

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
        # Get environment variables
        self.spLogin = config("SP_LOGIN")
        self.spPassword = config("SP_PASSWORD")
        self.spLink = config("SP_LINK")
        self.spSite = config("SP_SITE")
        self.spList = config("SP_LIST")
        self.mongoClient = config("MONGO_CLIENT")
        self.mongoDatabase = config("MONGO_DATABASE")
        self.mongoCollection = config("MONGO_COLLECTION")
        self.authSpCookie = None
        self.authSpSite = None
        self.authSpList = None

    def __repr__(self):
        return "<" + self.authSpList + ">"

    def auth(self):
        # Authentication
        try:
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
        except Exception:
            print("SharePoint authentication failed.", sys.exc_info())

    def get(self):
        # Get items from the Lists
        try:
            # Get the list of the site
            data = self.authSpList.GetListItems("All Items", fields=["ID", "Title"])

            for item in data:
                print(item)
        except Exception:
            print("Failed getting SharePoint Lists.", sys.exc_info())

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

    def delete(self):
        # Delete items
        try:
            # Ids to delete
            deleteData = ["23", "24"]

            print("Deleting items...")
            deleted = self.authSpList.UpdateListItems(data=deleteData, kind="Delete")
            if deleted:
                print("Successfully deleted items!")
        except Exception:
            print("SharePoint Lists delete failed.", sys.exc_info())

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
                {"Title": "company two"}, {"$set": {"Title": "company six"}}
            )
            items = collection.find({})
            for item in items:
                print(item)
            collection.delete_many({"Title": "mongo test"})
            print("MongoDB finished...")
        except Exception:
            print("Unable to connect to the mongo server.", sys.exc_info())


def main():
    try:
        # Options list
        parser = OptionParser(
            usage="Usage: python3.7 %prog [options]", add_help_option=True
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
            "-d",
            "--delete",
            action="store_true",
            dest="spDelete",
            default=False,
            help="delete items of Microsoft List",
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

        if not (True in vars(opts).values()):
            showHelp()
            return

        # Creates new SharePoint instance
        sharepoint = SharePoint()

        # Authentication
        sharepoint.auth()

        # Create new items
        if "-c" in args:
            sharepoint.create()

        # Update the list
        if "-u" in args:
            sharepoint.update()

        # Delete items
        if "-d" in args:
            sharepoint.delete()

        # MongoDB
        if "-m" in args:
            sharepoint.mongo()
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
