from datetime import datetime
import pathlib, shutil, logging, sys, os, socket

# This script can be run automatically through Windows scheduler.
# Use pyinstaller to convert into a standalone exe (no need to install python on all systems). 
# Syntax: pyinstaller updateDBF-0.3.py

# This does not work if local database is in Program Files (problems w permissions).

# Basic process:
# 1. Run script to close any instances of ClickPrint and/or LabelView.
# 2. Get date modified for server and local files (this assumes local file is in C:\LABELVIEW DB\Label Files\DATABASE).
# 3. Compare file modified dates. 
# 3a. If equal, do nothing.
# 3b. If server is newer, copy to local. 
# 3c. If local is newer, log error message.


# ******* FILE PATHS - CHANGE HERE *******
local_path = pathlib.Path(r'C:\LABELVIEW DB\Label Files\DATABASE')
srv_path = pathlib.Path(r'\\SERVER_NAME\SERVER_DIR\Whse Label File')





# Error logging setup
hostname = socket.gethostname()
user = os.getlogin() + '@' + hostname
logpathtemp = pathlib.Path(r'\\SERVER_NAME\SERVER_DIR\update_dbf_auto\Logs')
logpath = pathlib.Path(logpathtemp, hostname + '_DBFsync.log')
FORMAT = user + ' %(levelname)s: %(asctime)s %(message)s'
logging.basicConfig(filename=logpath, encoding='utf-8', level=logging.INFO, format=FORMAT)


# Check for valid directories
if local_path.exists() == False:
    logging.error("Couldn't find " + local_path + ", check file paths.")
    sys.exit("Network error")
if srv_path.exists() == False:
    # log error locally on Desktop
    logpathtemp = str(pathlib.Path.home() + r'\Desktop')
    logpathlocal = pathlib.Path(logpathtemp, hostname + '_DBFsync.log')
    logging.basicConfig(filename=logpathlocal, encoding='utf-8', level=logging.INFO, format=FORMAT)
    logging.error("Couldn't reach " + srv_path + ", check file paths.")
    sys.exit("Network error")


# Close any open ClickPrint or LabelView processes
os.system("taskkill /f /im  ClickPrint.exe") # Click Print program
os.system("taskkill /f /im  Lppa.exe") # Labeling Software
os.system("taskkill /f /im  LV.exe") # LabelView program
os.system("@echo Please wait, comparing files...")



# Variables for specific files
local_dbf = pathlib.Path(local_path / 'DATAFILE.DBF')
local_access = pathlib.Path(local_path / 'ACCESS_FILE.accdb')
srv_dbf = pathlib.Path(srv_path / 'DATAFILE.DBF')
srv_access = pathlib.Path(srv_path / 'ACCESS_FILE.accdb')





# Compare date modified - DBF
mtime_srv = datetime.fromtimestamp(srv_dbf.stat().st_mtime)
mtime_local = datetime.fromtimestamp(local_dbf.stat().st_mtime)

if mtime_srv == mtime_local:
    # Match - local is up to date
    logging.info("Local DBF is up to date.")
    os.system("@echo Local DBF is up to date.")
    # continue on to Access file check

elif mtime_srv > mtime_local:
    # Local is behind - update
    logging.debug("Server is ahead of local DBF.")
    
    # Check that server file is reachable! Should be but who knows
    if srv_dbf.exists():
        try:
            # Delete old local file
            local_dbf.unlink()
            logging.debug("Local DBF deleted.")
            # Copy from server to local dir
            shutil.copy2(srv_dbf, local_dbf)
            logging.info("DBF copied from server.")
            os.system("@echo DBF copied from server.")
        except PermissionError:
            logging.error("Permission error: try running as admin.")
            os.system("@echo Permission error, update will now exit.")
            sys.exit("Permission error")
        except:
            logging.error("Unknown error occurred in updating DBF.")
            os.system("@echo Unknown error, update will now exit.")
            sys.exit(9)
    else:
        logging.error("Couldn't reach " + srv_dbf + " after clearing initial check. Network may be unstable.")
        os.system("@echo Network error, update will now exit.")
        sys.exit("Network error")
    # continue on to Access file check
    
elif mtime_srv < mtime_local:
    # Local is ahead - replace with server file
    logging.warning("Local is ahead of server DBF. Most likely this was due to a previous manual update (drag and drop). Update script will now replace local file with network file.")
    
    # Check that server file is reachable! Should be but who knows
    if srv_dbf.exists():
        try:
            # Delete old local file
            local_dbf.unlink()
            logging.debug("Local DBF deleted.")
            # Copy from server to local dir
            shutil.copy2(srv_dbf, local_dbf)
            logging.info("DBF copied from server.")
            os.system("@echo DBF copied from server.")
        except PermissionError:
            logging.error("Permission error: try running as admin.")
            os.system("@echo Permission error, update will now exit.")
            sys.exit("Permission error")
        except:
            logging.error("Unknown error occurred in updating DBF.")
            os.system("@echo Unknown error, update will now exit.")
            sys.exit(9)
    else:
        logging.error("Couldn't reach " + srv_dbf + " after clearing initial check. Network may be unstable.")
        os.system("@echo Network error, update will now exit.")
        sys.exit("Network error")
    # continue on to Access file check
else:
    # how did you get here?
    logging.critical("Something is wrong. How did you get here? Check that you actually have files on SERVER_NAME and on the local machine.")
    os.system("@echo Critical error, update will now exit.")
    sys.exit(9)




os.system("@echo Please wait, comparing files...")

# Compare date modified - Access
mtime_srv = datetime.fromtimestamp(srv_access.stat().st_mtime)
mtime_local = datetime.fromtimestamp(local_access.stat().st_mtime)

if mtime_srv == mtime_local:
    # Match - local is up to date
    logging.info("Local Access file is up to date.")
    os.system("@echo Local Access file is up to date.")
    # done

elif mtime_srv > mtime_local:
    # Local is behind - update
    logging.debug("Server is ahead of local Access file.")
    
    # Check that server file is reachable! Should be but who knows
    if srv_access.exists():
        try: 
            # Delete old local file
            local_access.unlink()
            logging.debug("Local Access file deleted.")
            # Copy from server to local dir
            shutil.copy2(srv_access, local_access)
            logging.info("Access file copied from server.")
            os.system("@echo Access file copied from server.")
        except PermissionError:
            logging.error("Permission error: try running as admin.")
            os.system("@echo Permission error, update will now exit.")
            sys.exit("Permission error")
        except:
            logging.error("Unknown error occurred in updating Access file.")
            os.system("@echo Unknown error, update will now exit.")
            sys.exit(9)
    else:
        logging.error("Couldn't reach " + srv_access + " after clearing initial check. Network may be unstable.")
        os.system("@echo Network error, update will now exit.")
        sys.exit("Network error")
    # done
    
elif mtime_srv < mtime_local:
    # Local is ahead - update from server.
    # This is a special case, seems like the Access file is 'updated' when accessed by ClickPrint/LV.
    logging.info("Local Access file was ahead of server.")

    # Check that server file is reachable! Should be but who knows
    if srv_access.exists():
        try: 
            # Delete old local file
            local_access.unlink()
            logging.debug("Local Access file deleted.")
            # Copy from server to local dir
            shutil.copy2(srv_access, local_access)
            logging.info("Access file copied from server.")
            os.system("@echo Access file copied from server.")
        except PermissionError:
            logging.error("Permission error: try running as admin.")
            os.system("@echo Permission error, update will now exit.")
            sys.exit("Permission error")
        except:
            logging.error("Unknown error occurred in updating Access file.")
            os.system("@echo Unknown error, update will now exit.")
            sys.exit(9)
    else:
        logging.error("Couldn't reach " + srv_access + " after clearing initial check. Network may be unstable.")
        os.system("@echo Network error, update will now exit.")
        sys.exit("Network error")
    # done

else:
    # how did you get here?
    logging.critical("Something is wrong. How did you get here? Check that you actually have files on SERVER_NAME and on the local machine.")
    os.system("@echo Critical error, update will now exit.")
    sys.exit(9)


# success
logging.info("Database sync completed successfully.")
os.system('echo Database sync completed successfully.')
sys.exit(0)
