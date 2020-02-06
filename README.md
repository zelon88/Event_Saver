NAME: Infrastructure_Heartbeat.vbs
 

TYPE: Visual Basic Script


PRIMARY LANGUAGE: 
VBS


AUTHOR: Justin Grimes


ORIGINAL VERSION DATE: 5/31/2019


CURRENT VERSION DATE: 2/6/2020


VERSION: v1.0


DESCRIPTION: 
A short and dirty script to preserve Event logs in potentially compromised environments.


PURPOSE: 
To monitor the event logs and preserve specified ones for administrator review.


INSTALLATION INSTRUCTIONS: 
1. Copy the entire "es.vbs" file to a well-hidden location on the local machine (the machine to be monitored).

2. Add a scheduled task to run the script on a designated schedule. Every 10m is adequate.

3. Check the logs frequently to see if they have captured the desired events.

4. Modify the configuration variables in the beginning of the ex.vbs file to adjust log location and preserved events.

NOTES: 
