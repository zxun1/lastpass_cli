from termcolor import colored
import pandas as pd

def readEXCEL(wName, sName, pcCol, folderCol):
    # use pandas to read excel file
    # wName = excel file name
    # sName = sheet name
    # pcCol = column header of PC/device column
    # folderCol = column header of folder column PS: folders are created in Lastpass server to store PC/device passwords.

    machines = []

    df = pd.read_excel(wName, sheetname = sName, keep_default_na = False)

    for row in df.index:
        if not str(df.ix[row,pcCol]) is "":
            machines.append([df[pcCol][row], df[folderCol][row]])

    return machines

def readCSV(fName, pcCol, folderCol):
    # use pandas to read csv file
    # fName = csv file name
    # pcCol = column header of PC/device column
    # folderCol = column header of folder column PS: folders are created in Lastpass server to store PC/device passwords.

    machines = []

    df = pd.read_csv(fName, keep_default_na = False)

    for row in df.index:
        if not str(df.ix[row, pcCol]) is "":
            machines.append([df.ix[row, pcCol], df.ix[row, folderCol]])

    return machines

def readTxt(dName, fName):
    #dName = device name
    #fName = folder name

    machines = []

    if dName:
        machines.append([dName, fName])

    return machines

def genforall(user, machines):
    from subprocess import call #Run command with arguments
    from subprocess import check_output #Run command with arguments and return the result
    from os import devnull
    """
    Lastpass interaction
    """
    # Login
    # check_output will return "Success: Logged in as xxx@xxx.com (user account)" if the user has been authenticated successfully.
    login = check_output([" lpass login "+user], shell=True)

    # Compare the result with first 7 char
    if(login[:7] == "Success"):
    # iterate machines
        for machine in machines:
            print ("Generating password for " + machine[0]+"   (o_o)")
            # This can be tweaked to have more granularity in the folder structure if required
            # removed symbols due to unknown compatability
            call([" lpass share create "+machine[1]], shell=True, stdout=open(devnull, 'wb'))
            call([" lpass generate --no-symbols " + "'" + machine[1]  + "'" + "/"  + "'" + machine[0] + "'" + " 20"], shell=True, stdout=open(devnull, 'wb'))
        # Sync vault
        call([" lpass sync"], shell=True)
        print(colored('Success:','green')+' Passwords generated and Sync\'d for all '+str(len(machines))+' entries')

        # Logout and clean up local vault cache
        call([" lpass logout -f"], shell=True)
    else:
        print "Login failed"

def genlpassEXCEL(filename, sheetname, pcCol, folderCol, user):
    try:
        genforall(user, readEXCEL(filename, sheetname, pcCol, folderCol))
    except Exception as e:
        print e

def genlpassCSV(filename, pcCol, folderCol, user):
    try:
        genforall(user, readCSV(filename, pcCol, folderCol))
    except Exception as e:
        print e

def genlpasstxt(filename, foldername, user):
    try:
        genforall(user, readTxt(filename, foldername))
    except Exception as e:
        print e

def get(dName, user):
    # dName = device name
    from subprocess import check_output
    from os import devnull
    from StringIO import StringIO
    from os import devnull

    login = check_output([" lpass login "+user], shell=True)
    if(login[:7] == "Success"):
        # next 3 rows of code convert str into pandas dataframe
        raw_result = check_output([ "lpass export --fields=name,password"], shell=True) #raw_result datatype is str
        result = StringIO(raw_result)
        df = pd.read_csv(result, sep=",")
        if dName in df['name'].values:
            row = df[df['name']==dName]['password'].values
            pw = row[0]
        else:
            print "No such device"
            pw = ""

        print pw

        return pw
    else:
        print "Login failed"

