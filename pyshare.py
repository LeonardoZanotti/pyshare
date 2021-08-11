#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

import sys

from decouple import config
from shareplum import Office365, Site
from shareplum.site import Version

# Get environment variables
spLogin = config('SP_LOGIN')
spPassword = config('SP_PASSWORD')
spLink = config('SP_LINK')
spSite = config('SP_SITE')
spList = config('SP_LIST')


def main():
    try:
        # Authentication
        authcookie = Office365(
            spLink,
            username=spLogin,
            password=spPassword
        ).GetCookies()

        # Login in the SharePoint
        site = Site(
            f'{spLink}/sites/{spSite}', version=Version.v365, authcookie=authcookie
        )

        # Enter the site
        sp_list = site.List(spList)

        # Get the list of the site
        data = sp_list.GetListItems('All Items')

        for item in data:
            print(item)

        # New data to create
        newData = [
            { 'Title': 'flawless' },
            { 'Title': 'Expert' }
        ]

        # Create new items
        print('Creating items...')
        created = sp_list.UpdateListItems(data=newData, kind='New')
        if created:
            print('Creation successful!')

        # Data to update
        updateData = [
            { 'ID': '3', 'Title': 'Ash' },
            { 'ID': '4', 'Title': 'Update' },
            { 'ID': '5', 'Title': 'Update 3' }
        ]

        # Update the list
        print('Updating items...')
        updated = sp_list.UpdateListItems(data=updateData, kind='Update')
        if updated:
            print('Update successful!')
    except:
        print('Error:', sys.exc_info())


if __name__ == '__main__':
    main()
