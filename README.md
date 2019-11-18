NAME: Ransomware_Defender

TYPE: VBS Script

PRIMARY LANGUAGE: VBScript
 
AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 8/23/2019

CURRENT VERSION DATE: 11/18/2019

VERSION: v1.6


DESCRIPTION: An application for early warning about potential ransomware activity on a domain workstation. 
On first run, Ransomware_Defender creates "Perimiter Files" in strategic places on the local filesystem.
On subsequent runs, Ransomware_Defender will check that the perimiter files still exist.\
If perimiter files are found we compare them to the original perimiter file. 
If perimiter files are missing they are searched for. 
If perimiterFiles have been tampered with the workstation will emit a Log, an email notification, and shut down to prevent further damage.





PURPOSE: To detect malicious file operations early enough that they do not cause widespread damage to company equipment.




INSTALLATION INSTRUCTIONS: 
1. Install Ransomware_Defender into a subdirectory of your Network-wide scripts folder.
2. Open Ransomware_Defender.vbs with a text editor and configure the variables at the start of the script to match your environment.
3. Open sendmail.ini with a text editor and configure your email server settings.
4. Run the script automatically on domain workstations at machine startup or user logon with a GPO. Or both!
5. Run the script automatically with scheduled tasks at regular intervals.
6. To add additional locations for monitoring, add the full absolute path to the "perimiterFiles" array.




NOTES: 
1. This script MUST be run with administrative rights.
2. If this script is started in regular user mode, it will prompt for administrator elevation.
3. "Fake Sendmail for Windows" is required for this application to send notification emails. Per the "Fake Sendmail" license, the required binaries are provided.
4. To reinstall "Fake Sendmail for Windows" please visit  https://www.glob.com.au/sendmail/
5. Use absolute UNC paths for network addresses. DO NOT run this from a network drive letter. The restartAsAdmin() function will not work properly.
6. If using as a startup/logon script it is advised to NOT use a conditional that checks for the prescence of the script prior to running it. Doing so could result in a false negative if ransomware damages Ransomware_Defender before it can be run. Errors produced by such a condition would alert users that something was wrong.
7. You may get a single false positive the first or second time the script is run. It is reccomended to either comment out the ojbShell.run line in killWorkstation() for the first couple of runs to give the script a chance to get perimiter files settled.