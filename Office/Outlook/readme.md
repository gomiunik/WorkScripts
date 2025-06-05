## ListAllFromAddressesInOutlookInbox.ps1
PowerShell script that you can run as unprivileged user. It opens your default inbox and recursively from all subfolders extracts "from" email addresses and at the end lists them in alphabetical order.
How to run: in command prompt open the folder where script is located and run 
```ps
PowerShell.exe .\ListAllFromAddressesInOutlookInbox.ps1
```
If you want the list exported e.g. to a csv file just append >list.csv
```ps
PowerShell.exe .\ListAllFromAddressesInOutlookInbox.ps1 >list.csv
```
