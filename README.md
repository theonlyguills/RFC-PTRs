# RFC-PTRs
Rockcliffe Algonquin PTR conversion macros for Excel

## Installation

Go to https://github.com/theonlyguills/RFC-PTRs

You will see this page

![my repo](images/gotosite.png)

Click on the latest release on the right hand side. Download the zipped source code (first link)

![releases](images/clicklatest.png)

Extract the contents to C:\RFC

It will look like this on disk:

![my repo](images/extracttoC.png)

## Unblocking the files

Depending on your version of Windows, you might have to unblock the files because they came from the internet and Windows does not like that. To check, right click the files **RFCAddIn.xlam** and **Rockliffe_Report_Macro.xlsm**, one by one, and go to Properties. If there is an 'Unblock' option, make sure it is checked.

This will look like this:

![unblock](images/unblock.png)

## Setting up the addin in Excel
Open a blank workbook in Excel. Right click on the ribbon, which is the grey area where the Bold, Italic, etc buttons are. Click Customize Ribbon...

![customize](images/customize.png)

In there click Add-ins in the left section then the Go... button next to Manage Excel Add-ins.

![manage addins](images/clickgo.png)

Click Browse and browse to the C:\RFC folder then select the addin file. You will now see the addin in the list and it should have a checkmark next to it.

![manage addins](images/addonadded.png)

You will now have a new tab in the ribbon called RFC with 2 buttons.

![manage addins](images/addedtab.png)

