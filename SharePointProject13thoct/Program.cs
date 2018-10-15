using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SharePointProject
{
    class Program
    {
        static void Main(string[] args)
        {
            string userName = "venu.kalam@acuvate.com";
            Console.WriteLine("Enter your password.");
            SecureString password = GetPassword();
            bool IsError = true;

            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/VenuProject13102018/"))
            {

                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);

                try
                {
                    List list = clientContext.Web.Lists.GetByTitle("Documents");

                    clientContext.Load(list);
                    clientContext.ExecuteQuery();

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
                    camlQuery.FolderServerRelativeUrl = "/teams/VenuProject13102018/Shared%20Documents";

                    ListItemCollection listitems = list.GetItems(camlQuery);

                    clientContext.Load(listitems, items => items.Include(i => i["Title"]));
                    clientContext.ExecuteQuery();
                    for (int i = 0; i < listitems.Count; i++)
                    {
                        SP.ListItem itemOfInterest = listitems[i];
                        if (itemOfInterest["Title"] != null)
                        {
                            string fileName = itemOfInterest["Title"].ToString();
                            if (i == 0)
                            {

                                ReadExcelData(clientContext, fileName);
                            }
                        }
                    }
                    IsError = false;
                }
                catch (Exception e)
                {
                    IsError = true;
                    Console.WriteLine(e.Message);
                    Console.WriteLine("first catch block main");
                }
                finally
                {
                    if (IsError)
                    {

                    }
                }
                Console.ReadKey();

            }



        }


        private static void ReadExcelData(ClientContext clientContext, string fileName)
        {

            bool IsError = true;
            string strErrorMsg = string.Empty;
            const string lstDocName = "Documents";

            try
            {

                DataTable dataTable = new DataTable("ExcelDataTable");
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();

                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + "ExcelFile.xlsx";
                SP.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);

                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();

                clientContext.Load(file);

                clientContext.ExecuteQuery();

                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {

                    if (data != null)
                    {

                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument documnet = SpreadsheetDocument.Open(mStream, false))
                        {

                            WorkbookPart workbookpart = documnet.WorkbookPart;

                            IEnumerable<Sheet> sheets = documnet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                            string relationshipId = sheets.First().Id.Value;

                            WorksheetPart worksheetPart = (WorksheetPart)documnet.WorkbookPart.GetPartById(relationshipId);

                            Worksheet workSheet = worksheetPart.Worksheet;

                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                            IEnumerable<Row> rows = sheetData.Descendants<Row>();

                            foreach (Cell cell in rows.ElementAt(0))
                            {

                                string str = GetCellValue(clientContext, documnet, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {

                                if (row != null)
                                {

                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {

                                        dataRow[i] = GetCellValue(clientContext, documnet, row.Descendants<Cell>().ElementAt(i));
                                        
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);





                        }
                    }
                }
                UpdateSPList(clientContext, dataTable, fileName);
                IsError = false;
            }
            catch (Exception e)
            {
                IsError = true;
                Console.WriteLine(e.Message);
                Console.WriteLine("second catch block");
            }
            finally
            {
                if (IsError)
                {
                    //Logging
                }
            }


        }

        private static void UpdateSPList(ClientContext clientContext, DataTable dataTable, string fileName)
        {

            bool isError = true;
            string strErrorMsg = string.Empty;
            Int32 count = 0;
            const string lstName = "Files";
            const string lstColCreatedBy = "create";
            const string lstColType = "typeof";
            const string lstColSize = "Size";
            try
            {

                SP.List oList = clientContext.Web.Lists.GetByTitle(lstName);

                foreach (DataRow row in dataTable.Rows)
                {
                    if (count++ == 0)
                        continue;


                    string filee = row[0].ToString();

                    string filename = filee.Split('\\').Last();
                    System.IO.FileInfo filesize = new System.IO.FileInfo(row[0].ToString());

                    long size = filesize.Length;
                    string exten = filesize.Extension;
                    Type type = filesize.GetType();
                    if ((size / 1048576.0) > 0 && (size / 1048576.0) < 15)
                    {

                        var fileCreationInformation = new FileCreationInformation();
                        fileCreationInformation.Content = System.IO.File.ReadAllBytes(row[0].ToString());


                        fileCreationInformation.Url = filename;
                        Microsoft.SharePoint.Client.File file = oList.RootFolder.Files.Add(fileCreationInformation);

                        clientContext.Load(file);
                        var item = file.ListItemAllFields;

                        item[lstColCreatedBy] = row[1].ToString();

                        item[lstColType] = exten;
                        item[lstColSize] = filesize.Length;
                        item.Update();

                        clientContext.ExecuteQuery();


                    }

                }

                isError = false;
            }
            catch (Exception e)
            {
                isError = true;
                Console.WriteLine(e.Message);
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
        }


        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
            return value;
        }



        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;

            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}




//private static void UpdateSPList(ClientContext clientContext, DataTable dataTable, string fileName)
//{
//    bool isError = true;
//    string strErrorMsg = string.Empty;
//    Int32 count = 0;
//    const string lstName = "UploadedFiles";
//    const string lstColTitle = "Title";
//    const string lstColDepartment = "Department";
//    const string lstColCreatedBy = "CreatedBy";
//    const string lstColType = "Type";
//    const string lstColSize = "size";
//    try
//    {

//        SP.List oList = clientContext.Web.Lists.GetByTitle(lstName);

//        //FieldCollection fields = oList.Fields;
//        //clientContext.Load(fields);
//        //clientContext.ExecuteQuery();
//        //foreach(SP.Field field in fields)
//        //{
//        //    Console.WriteLine(field.Title);
//        //}

//        foreach (DataRow row in dataTable.Rows)
//        {

//            System.IO.FileInfo filesize = new System.IO.FileInfo(@"C:\Users\venu.kalam\Documents\SharePointProject\images\bkg-blu.jpg");

//            long size = filesize.Length;
//            string exten = filesize.Extension;
//            Type type = filesize.GetType();
//            if ((size / 1048576.0) > 0 && (size / 1048576.0) < 15)
//            {
//                //    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
//                //    ListItem oListItem = oList.AddItem(itemCreateInfo);
//                //oListItem.FieldValues[lstColTitle]= row[0].ToString();
//                //   oListItem[lstColTitle] = row[0].ToString();
//                //     oListItem.FieldValues[lstColDepartment] = row[1].ToString();
//                //  oListItem[lstColCreatedBy] = row[2];
//                //  oListItem[lstColType] = row[3];
//                //  oListItem[lstColSize] = row[2];

//                var fileCreationInformation = new FileCreationInformation();
//                fileCreationInformation.Content = System.IO.File.ReadAllBytes(@"C:\Users\venu.kalam\Documents\SharePointProject\images\bkg-blu.jpg");


//                fileCreationInformation.Url = "bkg-blu.jpg";
//                Microsoft.SharePoint.Client.File file = oList.RootFolder.Files.Add(fileCreationInformation);

//                clientContext.Load(file);
//                var item = file.ListItemAllFields;
//                item[lstColDepartment] = "images";
//                // item[lstColCreatedBy] = row[1].ToString();
//                // item[lstColType] = exten;
//                item[lstColSize] = filesize.Length;
//                item.Update();
//                clientContext.ExecuteQuery();
//                Console.WriteLine("last line");
//                count++;
//            }
//        }

//        isError = false;
//    }
//    catch (Exception e)
//    {
//        isError = true;
//        Console.WriteLine(e.Message);
//    }
//    finally
//    {
//        if (isError)
//        {
//            //Logging
//        }
//    }
//}