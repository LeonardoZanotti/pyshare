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

# Authentication
authcookie = Office365(spLink,
                       username=spLogin, password=spPassword).GetCookies()

# Login in the SharePoint
site = Site(spSite,
            authcookie=authcookie)

# Enter the site
sp_list = site.List(spList)

# Get the list of the site
data = sp_list.GetListItems('All Items')

for item in data:
    print(item)
