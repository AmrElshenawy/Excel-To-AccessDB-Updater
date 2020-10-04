# Excel To Microsoft Access DB Updater

This C# executable custom made application serves the purpose of updating a Microsoft Access Database with specific information for thousands of Excel files.
Within the DB, there are two columns missing information for thousands of records in the DB. Manually accessing every single Excel file and updating the corresponding field in the DB is a job that would take weeks to finish!

Within this program, the user selects the Database they wish to update and forms a connection to the DB. Upon confirmation, the user will be displayed with the total number of records to be updated. Once the update is initated by the user, an additional feature runs first which cleans up the targeted directory from any unused files that has no corressponding field in the DB. The removed files from the directory are taken to a differend directory in the scenario that they need to be recovered.

Once the purge has been completed, the program initiates the main run of the program by accessing the Excel files, reading the thought-after information and then updating the DB with that recorded information. It's worth of noting that the Excel files are in a special format that are only accessibly using a custom-made Addin. Multiple secondary classes/projects are therefore imported into my project to allow me to access the Excel files using the references libraries.

Once the update has been completed, the user is notified with the status. An export button is available to export the entire recorded run of the porgram in a text file to a default directory for record-keeping.

Furthermore, in the event the that removed files from the main directory needs to be moved back, the user has the option to click a button which would move back all files into the main directory. On the next run of the program, the purge would run again, cleaning the directory of all files that have no corresponding information in the DB.

The program runs on two new threads to handle the information transfer in the background while keeping the GUI responsive.
