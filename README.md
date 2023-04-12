# aad-guest-accounts-owner-stamp
This script tries to solve the issue that AAD guest accounts do not have easily accessible information about who created them. 
This information is only available in audit logs and there is no easy way how to attribute a specific AAD guest account to the internal user who invited them.

What this scripts does is that it reads recently created AAD guest accounts (external users) and audit log events related to inviting external users for the same time period (by default 1 day back). It then tries to pair these events based on time when they occurred. A configurable time difference tolerance is used here (by default 60 seconds). If a match is found the scripts writes a stamp in the format "UPN_of_creator;timestamp" into attribute employeeType of the AAD guest account.  
