#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

import sys

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
            deleteData = ["9"]

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
            client = pymongo.MongoClient(
                "mongodb://localhost:27017/?readPreference=primary&appname=MongoDB%20Compass&ssl=false"
            )
            print(client.server_info())
            print("Connected!")
            print("Running payload...")
            db = client["pyshare"]
            collection = db["companies"]
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
        args = sys.argv

        # Creates new SharePoint instance
        sharepoint = SharePoint()

        # Authentication
        if "-a" in args:
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
    except:
        print("Error:", sys.exc_info())


if __name__ == "__main__":
    main()
