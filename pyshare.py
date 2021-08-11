#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare


import sys

import pymongo
from decouple import config
from shareplum import Office365, Site
from shareplum.site import Version

# Get environment variables
spLogin = config("SP_LOGIN")
spPassword = config("SP_PASSWORD")
spLink = config("SP_LINK")
spSite = config("SP_SITE")
spList = config("SP_LIST")


def main():
    try:
        args = sys.argv

        # Authentication
        authcookie = Office365(
            spLink, username=spLogin, password=spPassword
        ).GetCookies()

        # Login in the SharePoint
        site = Site(
            f"{spLink}/sites/{spSite}", version=Version.v365, authcookie=authcookie
        )

        # Enter the site
        sp_list = site.List(spList)

        # Get the list of the site
        data = sp_list.GetListItems("All Items", fields=["ID", "Title"])

        for item in data:
            print(item)

        # Create new items
        if "-c" in args:
            # New data to create
            newData = [{"Title": "Bingo"}, {"Title": "Expertise"}]

            print("Creating items...")
            created = sp_list.UpdateListItems(data=newData, kind="New")
            if created:
                print("Successfully created items!")

        # Update the list
        if "-u" in args:
            # Data to update
            updateData = [
                {"ID": "11", "Title": "Belest"},
                {"ID": "12", "Title": "Update 4"},
            ]

            print("Updating items...")
            updated = sp_list.UpdateListItems(data=updateData, kind="Update")
            if updated:
                print("Successfully updated items!")

        # Delete items
        if "-d" in args:
            # Ids to delete
            deleteData = ["9"]

            print("Deleting items...")
            deleted = sp_list.UpdateListItems(data=deleteData, kind="Delete")
            if deleted:
                print("Successfully deleted items!")

        # MongoDB
        if "-m" in args:
            print("Executing MongoDB process...")
            client = pymongo.MongoClient(
                "mongodb://localhost:27017/?readPreference=primary&appname=MongoDB%20Compass&ssl=false"
            )
            db = client["pyshare"]
            collection = db["companies"]
            collection.insert_one({"Title": "mongo test"})
            items = collection.find({})
            for item in items:
                print(item)
            print("MongoDB finished...")
    except:
        print("Error:", sys.exc_info())


if __name__ == "__main__":
    main()
