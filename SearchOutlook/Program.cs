using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace SearchOutlook
{
    class Program
    {
        static int Main(string[] args)
        {

            if (args.Length == 0)
            {
                Console.WriteLine("Please enter a search word.");
                Console.WriteLine("Usage: SearchOutlook <search word>");
                return 1;
            }

            EnumerateFoldersInDefaultStore(args[0]);
            return 0;
        }

        static void EnumerateFoldersInDefaultStore(string search)
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root, search);
        }

        // Uses recursion to enumerate Outlook subfolders.
        static void EnumerateFolders(Outlook.Folder folder, string search)
        {
            Outlook.Folders childFolders = folder.Folders;
            Console.WriteLine(childFolders.Count);
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        // Write the folder path.
                        Console.WriteLine("Enumerating " + childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        EnumerateFolders(childFolder, search);
                    }
                }
            }
            Console.WriteLine("Checking in " + folder.FolderPath);
            IterateMessages(folder, search);
        }

        static void IterateMessages(Outlook.Folder folder, string search)
        {

            // Iterate through all items ("messages") in a folder
            var fi = folder.Items;
            if (fi != null)
            {

                try
                {
                    foreach (Object item in fi)
                    {
                        Outlook.MailItem mailitem = (Outlook.MailItem)item;

                        string body = mailitem.Body;

                        if (body.Contains(search))
                        {
                            Console.WriteLine("\n");
                            Console.WriteLine("############################################################################");
                            Console.WriteLine(body);
                            Console.WriteLine("\n");
                        }
                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine("An error occurred: '{0}'", e);
                }
            }


        }
    }
}
