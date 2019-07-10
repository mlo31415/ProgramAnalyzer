import os
import sys

#=============================================================================
# Print the text to a log file open by the main program
# If isError is set also print it to the error file.
def Log(text, isError=False):
    global g_logFile        # File for ordinary logging
    global g_logHeader      # Text to be printed before the 1st log entry

    LogInit()

    # If this is the first log entry for this header, print it and then clear it so it's not printed again
    if g_logHeader is not None:
        print(g_logHeader)
        print("\n"+g_logHeader, file=g_logFile)
    g_logHeader=None

    if isError:
        LogError(text)

    # Print the log entry itself
    print(text)
    print(text, file=g_logFile)
    if isError:
        print(text, file=g_logErrorFile)


#=============================================================================
# Print the text to a log file open by the main program
# If isError is set also print it to the error file.
def LogError(text):
    global g_logErrorFile       # File for error logging
    global g_logErrorFileName   # Name of the error file
    global g_logErrorHeader     # Text to be printed before the 1st error log entry
    global g_errorsLogged       # Number of errors logged

    LogInit()

    # If this and error entry and is the first error entry for this header, print it and then clear it so it's not printed again
    if g_logErrorHeader is not None:
        print("----\n"+g_logErrorHeader, file=g_logErrorFile)
    g_logErrorHeader=None

    # Print the log entry itself
    print(text)
    print(text, file=g_logErrorFile)
    g_errorsLogged+=1


#***************************************************************
# Initialize any globals that have not yet been initialized.
def LogInit():
    global g_logErrorFile  # File for error logging
    global g_logErrorFileName  # Name of the error file
    global g_logErrorHeader  # Text to be printed before the 1st error log entry
    global g_errorsLogged

    if 'g_logErrorFileName' not in globals():
        g_logErrorFileName="Error report.txt"
    if 'g_logErrorFile' not in globals():
        g_logErrorFile=open(g_logErrorFileName, "w+")
    if 'g_logErrorHeader' not in globals():
        g_logErrorHeader=None
    if 'g_errorsLogged' not in globals():
        g_errorsLogged=0

    global g_logFile  # File for error logging
    global g_logFileName  # Name of the error file
    global g_logHeader  # Text to be printed before the 1st error log entry

    if 'g_logFileName' not in globals():
        g_logFileName="Log.txt"
    if 'g_logFile' not in globals():
        g_logFile=open(g_logFileName, "w+")
    if 'g_logHeader' not in globals():
        g_logHeader=None

# Set the header for any subsequent log entries
# Note that this header will only be printed once, and then only if there has been a log entry
def LogSetNewHeader(name):
    global g_logHeader
    global g_logErrorHeader
    global g_logLastHeader

    LogInit()

    if g_logLastHeader is None or name != g_logLastHeader:
        g_logHeader=name
        g_logErrorHeader=name
        g_logLastHeader=name


def LogOpen(logfilename, errorfilename):
    LogInit()

    global g_logFile
    global g_logFileName
    g_logFileName=logfilename
    g_logFile=open(g_logFileName, "w+")

    global g_logErrorFile
    global g_logErrorFileName
    g_logErrorFileName=errorfilename
    g_logErrorFile=open(g_logErrorFileName, "w+")

    global g_logHeader
    g_logHeader=None
    global g_logErrorHeader
    g_logErrorHeader=None
    global g_logLastHeader
    g_logLastHeader=None



def LogClose():
    if 'g_logFile' in globals():
        global g_logFile
        g_logFile.close()
        del g_logFile

    if 'g_logErrorFile' in globals():
        global g_logErrorFile
        global g_errorsLogged
        global g_logErrorFileName

        if g_logErrorFile is not None:
            if 'g_errorsLogged' in globals() and g_errorsLogged > 0:
                if sys.platform == "win32":
                    g_logErrorFile.close()
                else:
                    i=0 # Need a call here for Linux, Max
                os.system("notepad.exe "+g_logErrorFileName)
        del g_logErrorFile