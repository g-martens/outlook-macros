# Inleiding
Deze macro's kan je binnen outlook toevoegen
Macros beheren kan je via alt + F8


# FindFolder
This example finds a folder by its name, and you can choose to get the found folder activated. You can either enter the full name, or use wildcards. '*' and '%' are allowed as wildcards. Upper/lower cases will be ignored. See the constant SpeedUp in the head of the module: The default setting of True means the search will be somewhat faster, however, Outlook will be blocked during the search. If the search takes too long due to a great many folders, you can set the value to False. Set the constant StopAtFirstMatch to False if the search should not stop with the first match.

Copy the code into the module 'ThisOutlookSession'. For instance, it can be started by pressing alt+f8.


