#!/usr/bin/env python3.7
# Leonardo Zanotti
# https://github.com/LeonardoZanotti/pyshare

from shareplum import Office365, Site

authcookie = Office365('https://abc.sharepoint.com',
                       username='username@abc.com', password='password').GetCookies()
site = Site('https://abc.sharepoint.com/sites/MySharePointSite/',
            authcookie=authcookie)
sp_list = site.List('list name')
data = sp_list.GetListItems('All Items', rowlimit=200)
for item in data:
    print(item)
