import urllib.request

def Auth(username, password, url):
    passman = urllib.request.HTTPPasswordMgrWithDefaultRealm()
    passman.add_password(None,url,username,password)

    authhandler = urllib.request.HTTPBasicAuthHandler(passman)

    opener = urllib.request.build_opener(authhandler)

    urllib.request.install_opener(opener)

    pagehandle = urllib.request.urlopen(url)
    return pagehandle