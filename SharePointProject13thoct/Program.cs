using CsvHelper;
using Microsoft.SharePoint.Client;
//using Microsoft.SharePoint.Client.Runtime;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using SP = Microsoft.SharePoint.Client;
using System.Data;

namespace SharePointProject13thoct
{
    class Program
    {
        static void Main(string[] args)
        {
            string userName = "venu.kalam@acuvate.com";
            Console.WriteLine("Enter your password.");
            SecureString password = GetPassword();
           
            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/VenuProject13102018/"))
            {
              
                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);



                SP.List oList = clientContext.Web.Lists.GetByTitle("Documents");
                //clientContext.Load(oList);
                //clientContext.ExecuteQuery();
                //Console.WriteLine(oList.Title,oList.Id);
                ReadFileName(clientContext);

               

                //readXLS("https://acuvatehyd.sharepoint.com/teams/VenuProject13102018/Shared%20Documents/Forms/.aspx");
                // List oList = clientContext.Web.Lists.GetByTitle("Documents");
                //   clientContext.Load(oList);
                //   clientContext.ExecuteQuery();
                //ListItem listItem = oList.GetItemById(1);
                //clientContext.Load(listItem);
                //clientContext.ExecuteQuery();
                //  Console.WriteLine(oList.Title);
                //  Console.WriteLine(listItem.Client_Title);
                Console.ReadKey();
            }
        }

        private static void ReadFileName(ClientContext clientContext)
        {

            
            string fileName = string.Empty;
            bool isError = true;
            const string fldTitle = "LinkFilename";
            const string lstDocName = "Documents";
            const string strFolderServerRelativeUrl = "teams/VenuProject13102018/Shared%20Documents";
            string strErrorMsg = string.Empty;
            try
            {
                
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
              
                camlQuery.FolderServerRelativeUrl = strFolderServerRelativeUrl;
                
                ListItemCollection listItems = list.GetItems(camlQuery);
                Console.WriteLine("beforwe for loop");
                clientContext.Load(listItems);
                clientContext.ExecuteQuery();
                
                for (int i = 0; i < listItems.Count; i++)
                {
                    Console.WriteLine("eneterd try ReadFileName");
                    Console.WriteLine("entered first for loop45");
                    SP.ListItem itemOfInterest = listItems[i];
                    if (itemOfInterest[fldTitle] != null)
                    {
                        fileName = itemOfInterest[fldTitle].ToString();
                        if (i == 0)
                        {
                            Console.WriteLine(fileName);
                            ReadExcelData(clientContext, itemOfInterest[fldTitle].ToString());
                        }
                    }
                }
                Console.WriteLine("out of for loop");
                isError = false;
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
        }



        private static void ReadExcelData(ClientContext clientContext, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            const string lstDocName = "Documents";
            try
            {
                DataTable dataTable = new DataTable("EmployeeExcelDataTable");
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }
                    }
                }
                UpdateSPList(clientContext, dataTable, fileName);
                isError = false;
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
        }

        private static void UpdateSPList(ClientContext clientContext, DataTable dataTable, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            Int32 count = 0;
            const string lstName = "FileDetails";
            const string lstColTitle = "Title";
            const string lstColAddress = "CreatedBy";
            try
            {
                string fileExtension = ".xlsx";
                string fileNameWithOutExtension = fileName.Substring(0, fileName.Length - fileExtension.Length);
                if (fileNameWithOutExtension.Trim() == lstName)
                {
                    SP.List oList = clientContext.Web.Lists.GetByTitle(fileNameWithOutExtension);
                    foreach (DataRow row in dataTable.Rows)
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem oListItem = oList.AddItem(itemCreateInfo);
                        oListItem[lstColTitle] = row[0];
                        oListItem[lstColAddress] = row[1];
                        oListItem.Update();
                        clientContext.ExecuteQuery();
                        count++;
                    }
                }
                else
                {
                    count = 0;
                }
                if (count == 0)
                {
                    Console.Write("Error: List: '" + fileNameWithOutExtension + "' is not found in SharePoint.");
                }
                isError = false;
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


        //public static void readXLS(string FilePath)
        //{
        //    FileInfo existingFile = new FileInfo(FilePath);
        //    using (ExcelPackage package = new ExcelPackage(existingFile))
        //    {
        //        //get the first worksheet in the workbook
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
        //        int colCount = worksheet.Dimension.End.Column;  //get Column Count
        //        int rowCount = worksheet.Dimension.End.Row;     //get row count
        //        for (int row = 1; row <= rowCount; row++)
        //        {
        //            for (int col = 1; col <= colCount; col++)
        //            {
        //                Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value.ToString().Trim());
        //            }
        //        }
        //    }
        //}

        //private static List GetRecordsFromCsv()
        //{
        //    List records = new List();
        //    using (var sr = new StreamReader("https://acuvatehyd.sharepoint.com/:x:/r/teams/VenuProject13102018/_layouts/15/Doc.aspx?sourcedoc=%7B17283a1b-a6d5-4224-98be-0db8717ee03b%7D&action=default&uid=%7B17283A1B-A6D5-4224-98BE-0DB8717EE03B%7D&ListItemId=3&ListId=%7BEE676EA5-70BC-45F8-8556-A74C93B7BB13%7D&odsp=1&env=prod")) ;

        //    {
        //        var reader = new CsvReader(sr);
        //        records = reader.GetRecords().ToList();
        //    }

        //    return records;
        //}

        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
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
