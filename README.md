![](https://download.logo.wine/logo/SharePoint/SharePoint-Logo.wine.png)

# PyShare
Python program to interact with the [Microsoft SharePoint](https://www.microsoft.com/en-ww/microsoft-365/sharepoint/collaboration), specifically the Microsoft Lists.

## Installation
First, make sure you have [Python3.7](https://www.python.org/) or higher installed. Also, make sure you have the [Pip](https://pypi.org/project/pip/) of that python version.

So, enter the project folder and type the following commands:

```shell
$ cp .env.example .env                        # create the .env file
$ pip3.7 install SharePlum python-decouple    # install dependencies (use your pip version)
```

Now, just fill the `.env` with your credentials from the SharePoint.
```properties
SP_LOGIN=<your-Microsoft-email>
SP_PASSWORD=<your-Microsoft-password>
SP_LINK=<https://abc.sharepoint.com>
SP_SITE=<MySharePointSite>                    # from https://abc.sharepoint.com/sites/MySharePointSite
SP_LIST=<list-name>
```

Then, just run the program:
```shell
$ python3.7 pyshare.py                        # use your python version
```

## References
[SharePlum Documentation](https://pypi.org/project/SharePlum/)