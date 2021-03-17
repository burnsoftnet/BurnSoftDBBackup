# BurnSoftDBBackup

The BurnSoftDBBackup is the Main make up program that I use to copy the database from the local drive to the selected destination.  This program is commonly used in the My Loaders Log and My Gun Collection application.

## How to Use

The bulk of the configuration is in the app.config file.

Example:

`<appSettings>`
    `<add key="AppName" value="DBBackup" />`
    `<add key="MainAppName" value="My Gun Collection" />`
    `<add key="DBName" value="MGC.mdb" />`
    `<add key="RegKey" value="Software\BurnSoft\BSMGC\" />`
    `<add key="CheckProcess" value="false" />`
    `<add key="LogFilename" value="dbbackup.err.log" />`
    `<add key="AppABV" value="MGC" />`
  `</appSettings>`
  
  This will read the Local User Registry Key settings that are created by the My Gun Collection or My Loaders Log Application.  Just put in the Path to those application in the **RegKey** settings for it to read the value of were the database is located at.  
  
  The **MainAppName** will be the title of the application, while the **DBName** will also have to be changed to match the database that you are wanting to copy/backup.
  
  ## Release Log
  
  ###v4.10.357.8765
  - Initial Production release, also start of Open Source Init