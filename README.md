# AlphaTR

AlphaTR is a VBA program designed to format and edit a transcriptionists work prior to it reaching the QA. The editing and formatting is all done through a series of macros that check the formatting to the official QA style guide.
---

### Install Instructions - (Windows)
1. Open Windows Explorer/File Exporer
2. Copy/Paste the following into the address bar:
    - `%Appdata%/Microsoft/word`
    - If there isn't already a folder named STARTUP, create one
3. Open the STARTUP folder
4. Copy AlphaTR.dotm file into the STARTUP folder.
5. Restart Word

### Install Instructions - (Mac)
1. Double-click on the AlphaTR.dotm file
2. When prompted, click Enable Macros
3. Click on the View tab
4. Click on Macros and then the Organizer button
5. On the left, where it says, "Macro Project Items available in:", click the dropdown menu and select AlphaTR.dotm
6. Click on each one and then click the copy button to transfer them to Normal.dotm
    - NOTE: if the file name changes for whateevr reason, click rename and change the file back to the original name displayed on the left
7. Click Close and restart Word

---

### Using Macros
On the View tabe, you'll see a Macros button. By clicking on the button, you'll see the available macros. You can double click or click run to use them.

---

### TR Checklist

![TR Checklist](https://i.imgur.com/JrhJHxql.png)

TR checklist currently has four basic features. Tag Count, Cleanup, and then direct links to both the Key Style Guide and the Terminology Database.
Tag Count
Tag Count will tell you the total number of inaudibles, guesses, crosstalks, foreign tags, and research tags in any file.
Cleanup
Cleanup is the primary function of TR Checklist. It will ASSIST in format the transcript to AlphaScribes requirements in order to deliver the cleanest transcript possible and aid in eliminating unnecessary point deductions later on. It is advised to not simply rely on the program to format the transcript, but rather use it as a tool. Cleanup performs the following tasks:
- Eliminates two spaces that start a sentence
- Replaces straight quotes with curly quotes
- Doesn’t change the user’s specified settings
- Searches for incorrectly formatted timestamps
- Searches for incorrect punctuation
- Runs bug fixes in the background to correct a common spacing issue that Word creates caused by incorrect borders 
- Runs Spellcheck
