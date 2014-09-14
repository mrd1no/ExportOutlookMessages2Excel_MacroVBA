ExportOutlookMessages2Excel_MacroVBA
====================================

A simple MS Outlook VBA macro designed to export selected messages into an excel file
This macro is already working,however there are still a few developments I would like to implement

#Getting started
Change the permissions in MS Outlook in the Trust Center:
  * open File > Options > Trust Center > Trust Center Settings.. > Macro Settings
  * choose the security level "Notification for all macros"
In order to open VBA in MS Outlook, click Alt+F11
  Create a new module

#About the code
Before to use it remember to:
  * insert your windows username in the code
  * create a file where the macro can export the data
  *assure that the sheet where you want to export the messages has the same name in the macro
Edit the code using what you need:
  * in VBA library you can find other criteria for Outlook.MailItem
  * I choose what I need: SenderName, SenderEmailAddress,Subject, To, ReceivedTime

  
  
#Additional developments
I would like to fix a few bugs / to add new features:
  * to insert headlines in the exported file in excel (a first raw with labels)
  * to export entire folders and not just selected files
  * to export all messages (archieves included) keeping track of categories and folders
