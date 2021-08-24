![](https://download.logo.wine/logo/SharePoint/SharePoint-Logo.wine.png)

# PyShare
Python program to interact with the [Microsoft SharePoint](https://www.microsoft.com/en-ww/microsoft-365/sharepoint/collaboration), specifically the Microsoft Lists.

## Installation
First, make sure you have [Python3.7](https://www.python.org/) or higher installed. Also, make sure you have the [Pip](https://pypi.org/project/pip/) of that python version.

So, enter the project folder and type the following commands:

```shell
$ cp .env.example .env                        # create the .env file
$ pip3.7 install SharePlum python-decouple pymongo    # install dependencies (use your pip version)
```

Now, just fill the `.env` with your credentials from the SharePoint.
```properties
SP_LOGIN=<your-Microsoft-email>
SP_PASSWORD=<your-Microsoft-password>
SP_LINK=<https://abc.sharepoint.com>
SP_SITE=<MySharePointSite>                    # from https://abc.sharepoint.com/sites/MySharePointSite
SP_LIST=<list-name>
MONGO_CLIENT=<mongo-connection-string>
MONGO_DATABASE=<database-name>
MONGO_COLLECTION=<collection-name>
```

Then, just run the program:
```shell
$ python3.7 pyshare.py                        # use your python version
```

## References
[SharePlum Documentation](https://pypi.org/project/SharePlum/)

[PyMongo Documentation](https://pymongo.readthedocs.io/en/stable/)

[OptParser Documentation](https://docs.python.org/3/library/optparse.html)

[Performing A CRUD Operation On A SharePoint List Using Python - Ashirwad Satapathi](https://www.c-sharpcorner.com/article/performing-crud-operation-on-sharepoint-list-using-python/)

[Alternative JS SharePoint API - SpRestLib](https://gitbrent.github.io/SpRestLib/)