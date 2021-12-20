# Mail Merge for GS
This is a Google Sheets Add On written in Google Apps Script for simple mail merge function.
## How to use?
### Method 1
1. Install this Add On in a new Google Sheet
2. Run "Initialise Template" - this would create a "Template" worksheet for the recipient list.
3. Run "Initialise Configurations" - this would create a "Configuration" worksheet required to successful run a mail merge operation.
4. (Optional) Download the email banner to your Google Drive. Get the File Link. Modify to your needs.
5. Download the email html template to your Google Drive. Get the File Link. Modify to your needs.
6. Modify the template and configurations worksheet according to your setup.
7. Run "Notify By Email" to send emails.

See below sections for details about a template worksheet and configuration worksheet. I have also provided the sample files for banner, template and attachment.

### Method 2
If you organisation does not allow you to install Add Ons, you can deploy and use this script by copying it to a local Google Apps Script project and deploy as an Add On for your own use.

## Configurations
You must have this work sheet before running the mail merge function. Running this function will create a configuration each time. Rename the previously created configuration before running this function as it will remove and recreate a new "Configuration" worksheet each time.

```
Key	Value
startRow	6
lastRowOffset	0
banner	https://drive.google.com/file/d/1q68Mk9RuBkMVnyLvX3C1t2ehcH24uGe4/view?usp=sharing
template	https://drive.google.com/file/d/18YueNqDdP1IE6m3WEDFMZG-mGme-iv8T/view?usp=sharing
fromName	APAC Enablement [Mail Merge for GS]
fromEmail	nicholas.cs.chia@gmail.com
ccEmails	nicholas.cs.chia@gmail.com, nicholas.cs.chia@gmail.com
attachment	https://drive.google.com/file/d/1ko3zWBy9Js15XaG0QmH86RryyfHrORRp/view?usp=sharing
```

startRow: First row of student data. Currently this is row 6

lastRowOffset: No offset is 0

banner: Banner for HTML Email

template: HTML Email template

fromName: Sender Name

fromEmail: Sender email address

ccEmails: CC email address comma-separated

attachment: Attachment file. Must be PDF.

## Initialise Template
Running this function will create a template name "Template" each time. Rename the previously created template before running this function as it will remove and recreate a new "Template" worksheet each time.

You can have multiple worksheets of recipient list in the same spreadsheet for each mail merge operation.

```
name	Advanced Deployment with Red Hat Ansible Automation		
startDate	January 17, 2022		
endDate	January 21, 2022		
trgSchedule	Mon-Thurs: 8am-4pm & 8am-11am		
First Name	Last Name	Email Address	Start Enablement Email Sent by Who & When?
Nicholas1	Chia	nicholas.cs.chia@gmail.com	Sent on Tue Dec 07 2021 16:52:21 GMT+0800 (Singapore Standard Time) by Mail Merge for GS.
Nicholas2	Chia	nicholas.cs.chia@gmail.com	Sent on Tue Dec 07 2021 16:52:21 GMT+0800 (Singapore Standard Time) by Mail Merge for GS.
Nicholas3	Chia	nicholas.cs.chia@gmail.com	Sent on Tue Dec 07 2021 16:52:21 GMT+0800 (Singapore Standard Time) by Mail Merge for GS.
Nicholas4	Chia	nicholas.cs.chia@gmail.com	Sent on Tue Dec 07 2021 16:52:21 GMT+0800 (Singapore Standard Time) by Mail Merge for GS.
```

name: Title of the course

startDate: Start date of the course

endDate: End date of the course

trgSchedule: The period of the course

First Name: First name of the student

Last Name: Last name if the student

Email Address: Student email address

Start Enablement Email Sent by Who & When?: This field must be empty the first time. Email will be sent if this cell is empty. The Add On will add this field after an email is sent.