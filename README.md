Export Outlook Messages To Excel (VBA Macro)
====================================

This is my first repository on Github and I want to share with you a simple <a href="http://bit.ly/msg2excel">MS Outlook VBA macro</a> designed to export selected messages into an excel file. <br>
This macro is already working, however there are still a few developments I would like to implement

#<b> Getting started </b><br><br>
Change the permissions in MS Outlook in the <i>Trust Center</i>:
  * open File > Options > Trust Center > Trust Center Settings.. > Macro Settings
  * choose the security level "Notification for all macros"
<br>
In order to open VBA in MS Outlook, click Alt+F11
  * Create a new module
<br><br>
#<b>About the code</b><br><br>
Before to use it remember to:
  * insert your <i>Windows username</i> in the code
  * <i>create a file</i> where the macro can export the data
  * assure that the <i>sheet</i> where you want to export the messages has the same name in the macro
<br><br>
Edit the code using what you need:<br>
  * in VBA library you can find other criteria for Outlook.MailItem
  * I choose what I need: SenderName, SenderEmailAddress,Subject, To, ReceivedTime
<br><br>
#Additional developments <br>
I would like to fix a few bugs / to add new features:
  * to insert headlines in the exported file in excel (a first raw with labels)
  * to export entire folders and not just selected files
  * to export all messages (archieves included) keeping track of categories and folders

<br><br> Feel you free to contact me on <a href="https://twitter.com/mauroberlanda">Twitter </a> or <a href="http://games4joke.wordpress.com/"> visit my blog</a>.
