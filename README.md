# SearchOutlook
A C# tool to search through a running instance of Outlook for keywords


Simply run the tool with the command "SearchOutlook [search word]" and it will search every email account and their subfolders in the running instance of Outlook, and print to the screen any email containing the keyword you searched for.
  
 Works well with execute-assembly of various C2 frameworks like Cobalt Strike.

Before compiling add the Microsoft Office 15.0 Object Library via Project->Add Reference->COM->Microsoft Office 15.0 Object Library
