#!/bin/sh
#
# University of Vermont
# Enterprise Technology Services
# 
# OS X Thunderbird-Exchange Migration
# Created ©2015 October 27 by Jonathan L. Trigaux
# Last modified 20151027.1631
#
# This script is intended to be bundled as a postflight/postinstall script to a standard OS X .pkg installer.
# It expects package-defined variables and will behave unexpectedly (or fail outright) if runby itself from the command line.
# 
# The script/package installer will:
# 1) Ask to quit any running instances of Thunderbird before continuing.
# 
# 2) Scan local user accounts for Thunderbird profiles.
# 
# 3) Convert any detected imap.uvm.edu configurations to Exchange-compatible settings.
# 
# 4) Re-open Thunderbird if it was open when the Installer was launched.
# 
# 5) On subsequent runs, provide the option to restore old Thunderbird settings, if any are found.
#
# Script must run as admin/root in order to get into all local user home directories.

# this part is needed to get some commands to work with OS X 10.5
export COMMAND_MODE=unix2003

#################################################################################################################################
### prep
# set installation volume (for running within package or from naked script)
INSTALLVOL="${3}"
if [[ "${INSTALLVOL}" == "" ]] ; 
then
	INSTALLVOL="/"
fi

# package resources path
RESOURCES=""
if [[ -d "${1}/Contents/Resources" ]] ; 
then
	# this is for bundle-style packages
	RESOURCES="${1}/Contents/Resources"
else
	# this is for flat packages
	RESOURCES=`"${INSTALLVOL}"bin/pwd -P`
fi

# system state
HARDWARE=`"${INSTALLVOL}"usr/sbin/system_profiler SPHardwareDataType | "${INSTALLVOL}"usr/bin/grep "Model Identifier:" | "${INSTALLVOL}"usr/bin/sed -E 's|.+: ||g'`
PLATFORM=`"${INSTALLVOL}"usr/sbin/system_profiler SPHardwareDataType | "${INSTALLVOL}"usr/bin/grep "Processor Name:" | "${INSTALLVOL}"usr/bin/sed -E 's|.+: ||g'`
COMPNAME=`"${INSTALLVOL}"usr/sbin/system_profiler SPSoftwareDataType | "${INSTALLVOL}"usr/bin/grep "Computer Name:" | "${INSTALLVOL}"usr/bin/sed -E 's|.+: ||g'`
OSVERSIONFULL=`""${INSTALLVOL}"usr/sbin/system_profiler" SPSoftwareDataType | "${INSTALLVOL}"usr/bin/grep "System Version: " | "${INSTALLVOL}"usr/bin/sed -E 's|[^0-9]+ ||g'`
OSVERSION=`"${INSTALLVOL}"bin/echo ${OSVERSIONFULL} | "${INSTALLVOL}"usr/bin/sed -E 's|10.([0-9]{1,2}).*|\1|g'`
RUNNER=`"${INSTALLVOL}"usr/bin/id -un`
HASGUI=`"${INSTALLVOL}"bin/ps -A | "${INSTALLVOL}"usr/bin/grep "[I]nstaller.app"`
TIMESTAMP=`"${INSTALLVOL}"bin/date -j`
### end prep
#################################################################################################################################

#################################################################################################################################
### functions
# if Installer wasn't called from the command line, bring app/window to the front, usually Finder or Installer
FOCUS ()
{
	if [[ "${HASGUI}" != "" ]] ; 
	then
		"${INSTALLVOL}"usr/bin/osascript -e "tell application \"${1}\" to activate" >> "${LOG}" 2>&1
	fi
}

# write to log file
LOGUPDATE ()
{
	"${INSTALLVOL}"bin/echo "${1}" >> "${LOG}"
}
### end functions
#################################################################################################################################

#################################################################################################################################
### log
# start log
"${INSTALLVOL}"bin/mkdir -pm 775 "${INSTALLVOL}Library/Logs/UVM"
LOG="${INSTALLVOL}Library/Logs/UVM/UVM_Thunderbird-Exchange_Migration.log"
	LOGUPDATE  "------------------------------------------------------------------"
	LOGUPDATE  "UVM Thunderbird/Exchange Migration, ${TIMESTAMP}"
	LOGUPDATE  "${HARDWARE} | ${PLATFORM} | OS X ${OSVERSIONFULL}"
	LOGUPDATE  "Installing to volume \"${INSTALLVOL}\" on ${COMPNAME} | Running as ${RUNNER}"
	LOGUPDATE  "------------------------------------------------------------------"
### end log
#################################################################################################################################

#################################################################################################################################
### run that baby!
# determine if Thunderbird is running, and if so, from what path
LOGUPDATE "Checking for running Thunderbird process..."
TBIRDOPEN=false
TBIRDPID=`"${INSTALLVOL}"bin/ps -A | "${INSTALLVOL}"usr/bin/grep [t]hunderbird | "${INSTALLVOL}"usr/bin/awk '{print $1}'`
TBIRDPATH=`"${INSTALLVOL}"bin/ps -A | "${INSTALLVOL}"usr/bin/grep [t]hunderbird | "${INSTALLVOL}"usr/bin/awk '{print $4}' | "${INSTALLVOL}"usr/bin/sed -E 's/(.*\.app).*/\1/'`

if [[ "${TBIRDPID}" != "" ]] ; 
then
	LOGUPDATE "Thunderbird is running."
	TBIRDOPEN=true
	if [[ "${HASGUI}" != "" ]] ; 
	then
		# running in a graphic user environment, so we'll prompt for user-affecting actions
		LOGUPDATE "Prompting to quit Thunderbird..."
		FOCUS Finder
		APPQUIT=`"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"Thunderbird must be closed to continue migration. OK to quit Thunderbird?\" buttons {\"Yes, quit Thunderbird.\", \"No, not yet.\"} default button 2 with title \"UVM Thunderbird/Exchange Migration\" with icon POSIX file (POSIX path of \"${RESOURCES}/tbird.icns\") giving up after 600" 2>&1`
		FOCUS Installer
		LOGUPDATE "User response: ${APPQUIT}"
			
		if [[ `"${INSTALLVOL}"bin/echo "${APPQUIT}" | "${INSTALLVOL}"usr/bin/grep "Yes"` != "" ]] ; 
		then
			LOGUPDATE "Stopping Thunderbird process..."
			"${INSTALLVOL}"bin/kill ${TBIRDPID} 2>&1 >> "${LOG}"
			# give it a few seconds to close
			"${INSTALLVOL}"bin/sleep 2
			
			# confirm Thunderbird actually quit
			TBIRDPID=`"${INSTALLVOL}"bin/ps -A | "${INSTALLVOL}"usr/bin/grep [t]hunderbird | "${INSTALLVOL}"usr/bin/awk '{print $1}'`
			if [[ "${TBIRDPID}" == "" ]] ; 
			then
				LOGUPDATE "Thunderbird successfully closed."
			else
				LOGUPDATE "Thunderbird process stop FAILED, notifying user and exiting installer."
				FOCUS Finder
				"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"Unable to quit Thunderbird. Please close Thunderbird and run this installer again.\" & return & return & \"Exchange migration of Thunderbird preferences has not occurred.\" & return & return & \"This installer will now exit.\" buttons {\"OK\"} default button 1 with title \"UVM Thunderbird/Exchange Migration\" with icon 0 giving up after 30"
				FOCUS Installer
				
				ENDTIME=`"${INSTALLVOL}"bin/date -j`
			
				LOGUPDATE "Installer exit at ${ENDTIME}"
				LOGUPDATE "------------------------------------------------------------------"
				LOGUPDATE ""
				LOGUPDATE ""
				exit 1
			fi
		else
			LOGUPDATE "User chose to not close Thunderbird (or dialog timed out), notifying user and exiting installer."
			FOCUS Finder
			"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"Thunderbird will not be closed. Please close Thunderbird and run this installer again.\" & return & return & \"Exchange migration of Thunderbird preferences has not occurred.\" & return & return & \"This installer will now exit.\" buttons {\"OK\"} default button 1 with title \"UVM Thunderbird/Exchange Migration\" with icon 0 giving up after 30"
			FOCUS Installer
			
			ENDTIME=`"${INSTALLVOL}"bin/date -j`
			
			LOGUPDATE "Installer exit at ${ENDTIME}"
			LOGUPDATE "------------------------------------------------------------------"
			LOGUPDATE ""
			LOGUPDATE ""
			exit 1
		fi
	else
		LOGUPDATE "Closing open Thunderbird."
		"${INSTALLVOL}"bin/kill ${TBIRDPID} 2>&1 >> "${LOG}"
		# give it a few seconds to close
		"${INSTALLVOL}"bin/sleep 2

		# confirm Thunderbird actually quit
		TBIRDPID=`"${INSTALLVOL}"bin/ps -A | "${INSTALLVOL}"usr/bin/grep [t]hunderbird | "${INSTALLVOL}"usr/bin/awk '{print $1}'`
		if [[ "${TBIRDPID}" == "" ]] ; 
		then
			LOGUPDATE "Thunderbird successfully closed."
		else
			LOGUPDATE "Thunderbird process stop FAILED, exiting installer."
			
			ENDTIME=`"${INSTALLVOL}"bin/date -j`
		
			LOGUPDATE "Installer exit at ${ENDTIME}"
			LOGUPDATE "------------------------------------------------------------------"
			LOGUPDATE ""
			LOGUPDATE ""
			exit 1
		fi
	fi
else
	LOGUPDATE "Thunderbird not detected as running during script execution, continuing..."
fi

# get list of local user accounts
LOGUPDATE "Getting local users list..."
USERLIST=`"${INSTALLVOL}"usr/bin/dscl . list /Users | "${INSTALLVOL}"usr/bin/grep -v -E "^[/_].*$"`
LOGUPDATE "Local user list is:"
LOGUPDATE "${USERLIST}"
USERCOUNT=`"${INSTALLVOL}"bin/echo ${USERLIST} | "${INSTALLVOL}"usr/bin/awk '{print NF}'`
LOGUPDATE "Local user count is: ${USERCOUNT}"

# iterate through local home directories for Thunderbird prefs
STATUS=""
for ((THISUSER=1; THISUSER <= ${USERCOUNT}; THISUSER++))
do
	MYUSER=`"${INSTALLVOL}"bin/echo ${USERLIST} | "${INSTALLVOL}"usr/bin/awk '{print $var}' var=${THISUSER}`
	MYHOME=`"${INSTALLVOL}"usr/bin/dscl . read "/Users/${MYUSER}" NFSHomeDirectory | "${INSTALLVOL}"usr/bin/awk '{print $2}'`

	LOGUPDATE "${THISUSER}. ${MYUSER}, Home directory detected as ${MYHOME}"

	LOGUPDATE "Scanning for Thunderbird profiles..."
		
	TBIRDPROFILES="${MYHOME}/Library/Thunderbird/Profiles"

	if [[ -d "${TBIRDPROFILES}" ]] ; 
	then
		for SALTDIR in `ls "${TBIRDPROFILES}"`
		do
			PREFS="${TBIRDPROFILES}/${SALTDIR}/prefs.js"
			TMPFILE="${PREFS}.UVM.Exchange.tmp"
			BACKUP="${PREFS}.UVM.Exchange.backup"
			
			# check for restore file
			if [[ -f "${BACKUP}" ]] ; 
			then
				LOGUPDATE "Previous migration backup file detected."
				if [[ "${HASGUI}" != "" ]] ; 
				then
					# running in a graphic user environment, so we'll prompt for user-affecting actions
					FOCUS Finder
					PREFSRESTORE=`"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"It appears Exchange migration has already occurred for profile:\" & return & return & \"${PREFS}/${SALTDIR}\" & return & return & \"Do you want to restore the pre-migration backup file or keep the current Exchange config?\" buttons {\"Restore old IMAP config from backup.\", \"Keep new Exchange config.\"} default button 2 with title \"UVM Thunderbird/Exchange Migration\" with icon 2 giving up after 600"`
					FOCUS Installer
					LOGUPDATE "User response: ${PREFSRESTORE}"

					if [[ `"${INSTALLVOL}"bin/echo "${PREFSRESTORE}" | "${INSTALLVOL}"usr/bin/grep "old"` != "" ]] ; 
					then
						LOGUPDATE "Restoring backup prefs.js file..."
						PREFSHASH=`/sbin/md5 "${BACKUP}" | /usr/sbin/awk '{print $4}'`
						"${INSTALLVOL}"bin/mv "${BACKUP}" "${PREFS}" 2>&1 >> "${LOG}"
						/usr/sbin/chown ${MYUSER}:staff "${PREFS}"
						/bin/chmod 700 "${PREFS}"
						
						# confirm restoration succeeded
						if [[ `/sbin/md5 "${PREFS}" | /usr/sbin/awk '{print $4}'` == "${PREFSHASH}" ]] ; 
						then
							LOGUPDATE "Restoration successful for ${PREFS}"
							FOCUS Finder
							"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"Old IMAP configuration restored to:\" & return & return & \"${PREFS}\" buttons \"OK\" default button 1 with title \"UVM Thunderbird/Exchange Migration\" with icon POSIX file (POSIX path of \"${RESOURCES}/tbird.icns\") giving up after 30"
							FOCUS Installer
						else
							LOGUPDATE "Restoration FAILED for ${PREFS}"
							FOCUS Finder
							"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"Attempt to restore old IMAP configuration to:\" & return & return & \"${PREFS}\" & return & return & \"has failed. Please contact the UVM TechTeam Helpline at helpline@uvm.edu or 802-656-2604 for assistance.\" & return & return & \"Installer will now continue scanning for remaining Thunderbird profiles.\" buttons \"OK\" default button 1 with title \"UVM Thunderbird/Exchange Migration\" with icon 0 giving up after 30"
							FOCUS Installer
						fi
					else
						LOGUPDATE "User selected to keep Exchange-migrated preferences, skipping restoration of old IMAP config."
					fi
				else
					LOGUPDATE "No GUI present, skipping restoration notification and leaving Exchange configuration in place."
				fi
			else
				if [[ -f "${PREFS}" ]] ; 
				then
					# make a backup copy
					LOGUPDATE "Creating backup copy of ${PREFS}"
					"${INSTALLVOL}"bin/cp "${PREFS}" "${BACKUP}" 2>&1 >> "${LOG}"

					# get server number for imap.uvm.edu
					for SERVERNUM in `"${INSTALLVOL}"usr/bin/grep "imap.uvm.edu" "${PREFS}" | "${INSTALLVOL}"usr/bin/grep "hostname" | "${INSTALLVOL}"usr/bin/awk -F "." '{print $3}'`
					do
						if [[ "${SERVERNUM}" != "" ]] ; 
						then
							# check for server_sub_directory key
							SERVERSUB=`"${INSTALLVOL}"usr/bin/grep "${SERVERNUM}.server_sub_directory" "${PREFS}"`

							if [[ "${SERVERSUB}" != "" ]] ; 
							then
								# remove server_sub_directory line
								LOGUPDATE "Removing '${SERVERNUM}.server_sub_directory' key from ${PREFS}"
								"${INSTALLVOL}"usr/bin/sed -i ".UVM.Exchange.tmp" -E 's/.*'"${SERVERNUM}"'\.server_sub_directory.*//' "${PREFS}"
					
								# confirm delete succeeded
								SUBDIR=`"${INSTALLVOL}"usr/bin/grep "${SERVERNUM}.server_sub_directory" "${PREFS}"`
					
								if [[ "${SUBDIR}" == "" ]] ; 
								then
									LOGUPDATE "Migration successful for ${SERVERNUM} in ${PREFS}"
									/bin/rm "${TMPFILE}"
									SERVERSUB=""
									SUBDIR=""
								else
									LOGUPDATE "Migration FAILED for ${PREFS}"
									if [[ "${HASGUI}" != "" ]] ; 
									then
										LOGUPDATE "Notifying user of failed migration and exiting installer"
										FOCUS Finder
										"${INSTALLVOL}"usr/bin/osascript -e "tell application \"Finder\" to display dialog \"Attempt to migrate:\" & return & return & \"${PREFS}\" & return & return & \"failed. Contact the UVM TechTeam Helpline at helpline@uvm.edu or 802-656-2604 for assistance.\" & return & return & \"Installer will now continue scanning for remaining Thunderbird profiles.\" buttons \"OK\" default button 1 with title \"UVM Thunderbird/Exchange Migration\" with icon 0 giving up after 30"
										FOCUS Installer
									fi
								fi
							else
								LOGUPDATE "No '${SERVERNUM}.server_sub_directory' key was found in ${PREFS}, nothing to modify."
							fi
						else
							LOGUPDATE "No further instances of 'imap.uvm.edu' was found in ${PREFS}"
						fi
					done
				else
					LOGUPDATE "No prefs.js file found in profile directory ${SALTDIR}"
				fi
			fi
		done
	else
		LOGUPDATE "No Thunderbird profiles detected for ${MYUSER}."
	fi
done


# re-open Thunderbird, if it was open at time of execution
if [[ ${TBIRDOPEN} == true ]] ; 
then
	LOGUPDATE "Re-opening Thunderbird application..."
	"${INSTALLVOL}"usr/bin/open "${TBIRDPATH}"
	FOCUS Installer
else
	LOGUPDATE "Thunderbird was not open when migration script was launched, leaving closed."
fi
### end run
#################################################################################################################################

# close out log
ENDTIME=`"${INSTALLVOL}"bin/date -j`
	LOGUPDATE "------------------------------------------------------------------"
	LOGUPDATE "UVM Thunderbird/Exchange Migration finished at ${ENDTIME}"
	LOGUPDATE "------------------------------------------------------------------"
	LOGUPDATE ""
	LOGUPDATE ""

exit 0