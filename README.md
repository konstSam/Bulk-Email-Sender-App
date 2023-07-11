# Bulk Email Sender App

1. Create a list in a .txt file, containing the name and email of each client per line {NAME}, {EMAIL}. (e.g. Sam Kostas, k.s@mail.com)

2. The pdf files of each person MUST called "{NAME}'s Portfolio".

3. The chart of each month MUST always called "Monthly_Chart.jpeg".

4. All these files, all pdf's for every client and the chart, MUST be in the SAME folder.

5. The location should be the same as in the script.

6. RUN the email_automationGUI.py


Error handling:
1. If Image file is not found, sends emails without Monthly Chart.
2. If Attachment is not found for someone, it sends email to this recipient w/o the pdf.
3. When user email credentials are not correct or accepted, an error message will be displayed "Failed to authenticate".
4. When clients.txt is not found, an error message is displayed.

