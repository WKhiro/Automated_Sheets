# Automated Sheets Work

My first experience with Selenium Python.

This is used to automate some work I do every week for some extra cash. It saves me a lot of copy pasting and mouse movement so I can rest my hands. It basically does this:

- Authorizes access to a Google Sheet file with a creds file on my local
- Opens a bunch of tabs (I can choose how many I want at once)
- Goes through each tab and waits for my inputs
- Based on my inputs, it automatically pastes a result into a spreadsheet of my choice, closes the current tab, and swaps to the next one

To get access to a creds file and enable google drive compatibility, I had to go to Google Cloud Console, create a new project, make credentials and generate a key for the project.

In the creds file itself, it shows the scope parameters, or the two libraries that I had to enable (in Google Cloud Console on the project) to gain access to Google Drive and Sheets.

In the creds file, there is also an email that needs to be shared as an editor with the file you need to access.
