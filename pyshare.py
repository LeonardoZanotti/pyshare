#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

from decouple import config
from shareplum import Office365, Site

# Get environment variables
spLogin = config('SP_LOGIN')
spPassword = config('SP_PASSWORD')
spLink = config('SP_LINK')
spSite = config('SP_SITE')
spList = config('SP_LIST')

print(spLogin, spPassword, spLink, spSite, spList)

# Authentication
authcookie = Office365('https://abc.sharepoint.com',
                       username='username@abc.com', password='password').GetCookies()

# Login in the SharePoint
site = Site('https://abc.sharepoint.com/sites/MySharePointSite/',
            authcookie=authcookie)

# Enter the site
sp_list = site.List('list name')

# Get the list of the site
data = sp_list.GetListItems('All Items', rowlimit=200)

for item in data:
    print(item)
