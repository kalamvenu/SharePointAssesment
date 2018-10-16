using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointProject
{
    // class for exception logging
    public static class ExceptionLogging
    {

        private static String ErrorlineNo, Errormsg, extype, ErrorLocation;

        // method to store the details in a text file
        public static void SendErrorToText(Exception ex)
        {
            var line = Environment.NewLine + Environment.NewLine;

            ErrorlineNo = ex.StackTrace.Substring(ex.StackTrace.Length - 7, 7);
            Errormsg = ex.GetType().Name.ToString();
            extype = ex.GetType().ToString();

            ErrorLocation = ex.Message.ToString();

            try
            {
                string filepath = @"C:\Users\venu.kalam\Documents\SharePointProject\ExceptionsLog\";  //Text File Path

                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);

                }
                filepath = filepath + DateTime.Today.ToString("dd-MM-yy") + ".txt";   //Text File Name Same as Date
                if (!System.IO.File.Exists(filepath))
                {
                    System.IO.File.Create(filepath).Dispose();
                }
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    string error = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Error Line No :" + " " + ErrorlineNo + line + "Error Message:" + " " + Errormsg + line + "Exception Type:" + " " + extype + line + "Error Location :" + " " + ErrorLocation + line;
                    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
                    sw.WriteLine("-------------------------------------------------------------------------------------");
                    sw.WriteLine(line);
                    sw.WriteLine(error);
                    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    sw.WriteLine(line);
                    sw.Flush();
                    sw.Close();

                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }


   public static class Constants
    {
       public static string URLofSite = "https://acuvatehyd.sharepoint.com/teams/VenuProject13102018/";

       public static string ServerRelativeURL = "/teams/VenuProject13102018/Shared%20Documents";

        public static string DocumentLibraryWhereExcelFileisPresent = "Documents";

        public static string NameofColumnWhereExcelFileisPresent = "Title";

        public static string PathToStoreExcelFile = @"D:\New folder\excelfilefolder";

    }

    public static class Fields
    {
       public const string lstName = "Files";
        public const string lstColCreatedBy = "create";
        public const string lstColType = "typeof";
        public const string lstColSize = "Size";
        public const string lstcolStatus = "Status";
        public const string lstcolDepartment = "Department";
    }

    static class Settings
    {

    }
}
