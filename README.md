This project can be used to analyze the response time of your team member in Outlook. 
It:
1. Lets you pick an Outlook mailbox folder, collects all emails from the last 2 weeks, then groups them by conversation ID
2. For each conversation, it finds the earliest team member reply and the latest customer/user email before that reply, then calculates the response time in minutes
3. Generates an HTML report with color-coded response times (green for ≤1hr, orange for 1-4hrs, red for >4hrs or no reply), along with aggregate stats like average response time
4. Either sends that report to admin@abc.com or opens it as a draft, depending on the AUTO_SEND_REPORT flag
5. It matches team members against a hardcoded sender list using substring matching, and includes diagnostic utilities for troubleshooting folder contents and date ranges.

