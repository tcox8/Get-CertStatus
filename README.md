# Get-CertStatus
This script will query the Certificate Authority and pull a list of certifiates based on common templates. It gets the expiration date/time and publishes it to a webpage showing the status. The script also sends out an email with counts of the number of certs expiring in 15 days, 30 days and 60 days. This script should be setup to run as a scheduled task.

![Table Example](Example Images/webpageExample.PNG?raw=true)

The certificate status shows up with a colored exclamation mark based on if the cert is expiring in 15 days (red), 30 days (yellow/orange), 60 days (blue), 61+ days(green). You can mouse over the exclamation mark to get a read of the status.

# Things to Edit to Make This Work For You
Edit the variables under "Variables to Edit" in teh script. 
