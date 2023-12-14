# IntuneChromeBookmarks
Script that formats your URL list so it's ready to be used to deploy Chrome bookmarks through Intune

Download Book1.xlsx to use as your bookmark template.

Fill out the bookmark template and save it to your C:\ drive without changing the name. 
![image](https://github.com/CyberSkyler/IntuneChromeBookmarks/assets/153866716/7a19faca-d172-4f55-a9d7-ef4611d463c0)

Fill out your template and run the PowerShell command. You will find the output of the script at C:\*BookmarkFolderName*.txt

Your output should look something like this:

[
  {
    "toplevel_name": "Test Name"
  },
  {
    "url": "URL 1",
    "name": "Website 1"
  },
  {
    "url": "URL 2",
    "name": "Website 2"
  },
]

