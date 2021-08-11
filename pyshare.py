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
            try:
                # New data to create
                newData = [{"Title": "Bingo"}, {"Title": "Expertise"}]

                print("Creating items...")
                created = sp_list.UpdateListItems(data=newData, kind="New")
                if created:
                    print("Successfully created items!")
            except Exception:
                print("SharePoint Lists creation failed.")

        # Update the list
        if "-u" in args:
            try:
                # Data to update
                updateData = [
                    {"ID": "11", "Title": "Belest"},
                    {"ID": "12", "Title": "Update 4"},
                ]

                print("Updating items...")
                updated = sp_list.UpdateListItems(data=updateData, kind="Update")
                if updated:
                    print("Successfully updated items!")
            except Exception:
                print("SharePoint Lists update failed.")

        # Delete items
        if "-d" in args:
            try:
                # Ids to delete
                deleteData = ["9"]

                print("Deleting items...")
                deleted = sp_list.UpdateListItems(data=deleteData, kind="Delete")
                if deleted:
                    print("Successfully deleted items!")
            except Exception:
                print("SharePoint Lists delete failed.")

        # MongoDB
        if "-m" in args:
            try:
                print("Connecting to MongoDB...")
                client = pymongo.MongoClient(
                    "mongodb://localhost:2707/?readPreference=primary&appname=MongoDB%20Compass&ssl=false"
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
                print("Unable to connect to the mongo server.")
    except:
        print("Error:", sys.exc_info())


if __name__ == "__main__":
    main()
