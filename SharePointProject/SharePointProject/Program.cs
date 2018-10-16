using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;



namespace SharePointProject
{



    class Program
    {
        static void Main(string[] args)
        {
            // user details
            Console.WriteLine("Enter the Username");
            string userName = Console.ReadLine();
            Console.WriteLine("Enter your Password.");
            SecureString passWord = GetPassword();
            bool IsError = true;
            // hard coded some values because i am  doing some changes
            using (var clientContext = new ClientContext(Constants.URLofSite))
            {

                clientContext.Credentials = new SharePointOnlineCredentials(userName, passWord);

                try
                {
                    
                    // getting the excel file located in the documents library.

                    List list = clientContext.Web.Lists.GetByTitle(Constants.DocumentLibraryWhereExcelFileisPresent);
                    clientContext.Load(list);
                    clientContext.ExecuteQuery();
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
                    camlQuery.FolderServerRelativeUrl = Constants.ServerRelativeURL;
                    ListItemCollection ListItems = list.GetItems(camlQuery);
                    clientContext.Load(ListItems, Items => Items.Include(i => i[Constants.NameofColumnWhereExcelFileisPresent]));
                    clientContext.ExecuteQuery();
                    for (int i = 0; i < ListItems.Count; i++)
                    {
                        SP.ListItem ExcelItem = ListItems[i];
                        if (ExcelItem[Constants.NameofColumnWhereExcelFileisPresent] != null)
                        {
                            string ExcelFileName = ExcelItem[Constants.NameofColumnWhereExcelFileisPresent].ToString();
                            if (i == 0)
                            {
                                ReadExcelData(clientContext, ExcelFileName); // will enter the method for the specific file
                            }
                        }
                    }
                    IsError = false;
                }
                catch (Exception e)
                {
                    IsError = true;
                    Console.WriteLine(e.Message);
                    ExceptionLogging.SendErrorToText(e);// storing the exception details in a text file
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

        // method to read the data from excel file and copy it into data table
        private static void ReadExcelData(ClientContext clientContext, string ExcelFileName)
        {

            bool IsError = true;
            string strErrorMsg = string.Empty;
          

            try
            {
                DataTable dataTable = new DataTable("ExcelDataTable");
                List list = clientContext.Web.Lists.GetByTitle(Constants.DocumentLibraryWhereExcelFileisPresent);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();

                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + "ExcelFile.xlsx";
                SP.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl); 

                ClientResult<System.IO.Stream> Data = file.OpenBinaryStream();// opening the excel file

                clientContext.Load(file);

                clientContext.ExecuteQuery();
                // using memorystream to read the data
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {

                    if (Data != null)
                    {

                        Data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument Document = SpreadsheetDocument.Open(mStream, false))
                        {

                            WorkbookPart workBookPart = Document.WorkbookPart;

                            IEnumerable<Sheet> Sheets = Document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                            string relationshipId = Sheets.First().Id.Value;

                            WorksheetPart worksheetPart = (WorksheetPart)Document.WorkbookPart.GetPartById(relationshipId);

                            Worksheet workSheet = worksheetPart.Worksheet;

                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                            IEnumerable<Row> rows = sheetData.Descendants<Row>();

                            foreach (Cell cell in rows.ElementAt(0))
                            {

                                string str = GetCellValue(clientContext, Document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {

                                if (row != null)
                                {

                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {

                                        dataRow[i] = GetCellValue(clientContext, Document, row.Descendants<Cell>().ElementAt(i));

                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);//removing the first row since it is empty 

                        }
                    }

                }
                UpdateSPList(clientContext, dataTable, ExcelFileName);

                IsError = false;
            }
            catch (Exception e)
            {
                IsError = true;
                Console.WriteLine(e.Message);
                ExceptionLogging.SendErrorToText(e);

            }
            finally
            {
                if (IsError)
                {
                    
                }
            }
        }


        // method to upload files to document library
        private static void UpdateSPList(ClientContext clientContext, DataTable dataTable, string fileName)
        {
            string currentPath = "";
            bool isError = true;
            string strErrorMsg = string.Empty;
            Int32 count = 0;
           
           try
            {
                                             
                List list = clientContext.Web.Lists.GetByTitle(Constants.DocumentLibraryWhereExcelFileisPresent);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();

                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + "ExcelFile.xlsx";
                SP.File fileSP = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);

                ClientResult<System.IO.Stream> data = fileSP.OpenBinaryStream();

                clientContext.Load(fileSP);

                clientContext.ExecuteQuery();

                SP.List oList = clientContext.Web.Lists.GetByTitle(Fields.lstName);

                foreach (DataRow row in dataTable.Rows) // updating document library based on the data in datatable
                {

                    try
                    {
                        if (count++ == 0)
                            continue;

                        currentPath = row[0].ToString();

                        string filee = row[0].ToString();

                        string filename = filee.Split('\\').Last();
                        System.IO.FileInfo filesize = new System.IO.FileInfo(row[0].ToString());

                        long size = filesize.Length;
                        string exten = filesize.Extension;
                        Type type = filesize.GetType();
                        if ((size / 1048576.0) > 0 && (size / 1048576.0) < 15) //checking whether the file is within the sizelimit or not
                        {

                            var fileCreationInformation = new FileCreationInformation();
                            fileCreationInformation.Content = System.IO.File.ReadAllBytes(row[0].ToString());


                            fileCreationInformation.Url = filename;
                            Microsoft.SharePoint.Client.File file = oList.RootFolder.Files.Add(fileCreationInformation);

                            clientContext.Load(file);
                            var item = file.ListItemAllFields;

                            item[Fields.lstColCreatedBy] = row[1].ToString();

                            item[Fields.lstColType] = exten;
                            item[Fields.lstColSize] = filesize.Length;
                            //  item[lstcolStatus] = row[3].ToString();
                            string TemporaryString = row[3].ToString();
                            string[] Multiple = TemporaryString.Split(',');
                            item[Fields.lstcolStatus] = Multiple;
                            item[Fields.lstcolDepartment] = row[2].ToString();
                            item.Update();

                            clientContext.ExecuteQuery();

                            dataTable = DataTableUpdated(currentPath, "NA", "Success", dataTable); // updated data table
                        }

                        else
                        {
                            throw new Exception("file size is out of range");
                        }

                    }
                    catch (Exception e)
                    {
                        dataTable = DataTableUpdated(currentPath, e.Message, "Failed", dataTable);
                        Console.WriteLine(e.Message);
                        ExceptionLogging.SendErrorToText(e);
                    }
                }

                isError = false;
            }
            catch (Exception e)
            {
                isError = true;
                Console.WriteLine(e.Message);
                ExceptionLogging.SendErrorToText(e);
                dataTable = DataTableUpdated(currentPath, e.Message, "Failed", dataTable);
            }
            finally
            {
                if (isError)
                {
                  
                }
            }

            string FileLocation = DataTableToExcel(dataTable);// converting data table to excel file

            UploadUpdatedExcelFile(clientContext, FileLocation);//uploading the excel file
        }


        // helper method to get the values from each cell in excel sheet
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
                
                ExceptionLogging.SendErrorToText(e);
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

        //method to convert data table to excel sheet
        static string DataTableToExcel(DataTable dataTable)
        {
            DataTable Table = dataTable.Copy();
            string ExcelFilePath = Constants.PathToStoreExcelFile; //path
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet workSheet = excelApp.ActiveSheet;
                for (int i = 0; i < Table.Columns.Count; i++)
                {
                    workSheet.Cells[1, (i + 1)] = Table.Columns[i].ColumnName;
                }
                for (int i = 0; i < Table.Rows.Count; i++)
                {
                    for (int j = 0; j < Table.Columns.Count; j++)
                    {
                        workSheet.Cells[(i + 2), (j + 1)] = Table.Rows[i][j];
                    }
                }
                Console.WriteLine();
                Console.WriteLine();
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(ExcelFilePath + ".xlsx");
                if (fileInfo.Exists)
                    fileInfo.Delete();
                workSheet.SaveAs(ExcelFilePath);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendErrorToText(ex);
                throw new Exception("ExportToExcel: \n" + ex.Message);
               
            }
            return ExcelFilePath;
        }


        // method to upload the updated excel file along with reason and status
        static void UploadUpdatedExcelFile(ClientContext clientContext, String FileLocation)
        {
            try
            {

                List list = clientContext.Web.Lists.GetByTitle(Constants.DocumentLibraryWhereExcelFileisPresent);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                var fileCreationInformation = new FileCreationInformation();
                fileCreationInformation.Content = System.IO.File.ReadAllBytes(FileLocation + ".xlsx");
                string filee = FileLocation;
                string filename = filee.Split('\\').Last();
                fileCreationInformation.Url = filename + ".xlsx";
                fileCreationInformation.Overwrite = true;
                Microsoft.SharePoint.Client.File Excelfile = list.RootFolder.Files.Add(fileCreationInformation);
                clientContext.Load(Excelfile);
                clientContext.ExecuteQuery();

            }
            catch (Exception ex)
            {
                ExceptionLogging.SendErrorToText(ex);
                throw new Exception(ex.Message);
               
            }

        }

        // method to update the data table along with reason and status
        static DataTable DataTableUpdated(String FilePath, String Reason, String Status, DataTable dataTable)
        {
            DataTable Datatable = dataTable;
            foreach (DataRow Datarow in Datatable.Rows)
            {
                if (Datarow[0].Equals(FilePath))
                {
                    Datarow[4] = Status;
                    Datarow[5] = Reason;
                }
            }
            return Datatable;
        }



        //public static void UpdateExcelUsingOpenXMLSDK(System.IO.MemoryStream fileName)
        //{
        //    // Open the document for editing.
        //    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
        //    {
        //        // Access the main Workbook part, which contains all references.
        //        WorkbookPart workbookPart = spreadSheet.WorkbookPart;
        //        // get sheet by name
        //        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Sheet1").FirstOrDefault();
        //        // get worksheetpart by sheet id
        //        WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
        //        // The SheetData object will contain all the data.workbookPart.Workbook.GetFirstChild<Sheets>();
        //     //   SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        //        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        //        Cell cell = GetCell(worksheetPart.Worksheet, "E", 4);
        //        cell.CellValue = new CellValue("hello");
        //        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        //        // Save the worksheet.
        //        worksheetPart.Worksheet.Save();
        //        // for recacluation of formula
        //        spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
        //        spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
        //    }
        //}
        //private static Cell GetCell(Worksheet worksheet,
        //string columnName, uint rowIndex)
        //{
        //    Row row = GetRow(worksheet, rowIndex);
        //    if (row == null) return null;
        //    var FirstRow = row.Elements<Cell>().Where(c => string.Compare
        //    (c.CellReference.Value, columnName +
        //    rowIndex, true) == 0).FirstOrDefault();
        //    if (FirstRow == null) return null;
        //    return FirstRow;
        //}
        //private static Row GetRow(Worksheet worksheet, uint rowIndex)
        //{
        //    Row row = worksheet.GetFirstChild<SheetData>().
        //    Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        //    if (row == null)
        //    {
        //        throw new ArgumentException(String.Format("No row with index {0} found in spreadsheet", rowIndex));
        //    }
        //    return row;
        //}








        //private static void writetoexcel(ClientContext clientContext, DataTable dataTable, string fileName)
        //{
        //    const string lstDocName = "Documents";
        //    List list = clientContext.Web.Lists.GetByTitle(lstDocName);
        //    clientContext.Load(list.RootFolder);
        //    clientContext.ExecuteQuery();
        //    string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + "ExcelFile.xlsx";
        //    SP.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
        //    ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
        //    clientContext.Load(file);
        //    clientContext.ExecuteQuery();
        //    using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
        //    {
        //        if (data != null)
        //        {
        //            data.Value.CopyTo(mStream);
        //            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(mStream, true))
        //            {
        //                SharedStringTablePart shareStringPart;
        //                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        //                {
        //                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        //                }
        //                else
        //                {
        //                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
        //                }
        //                int index = InsertSharedStringItem("HELLO", shareStringPart);
        //                // Insert a new worksheet.
        //                WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
        //                // Insert cell A1 into the new worksheet.
        //                Cell cell = InsertCellInWorksheet("D", 1, worksheetPart);    
        //                cell.CellValue = new CellValue(index.ToString());
        //                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        //                // Save the new worksheet.
        //                worksheetPart.Worksheet.Save();
        //                file.Update();
        //                clientContext.ExecuteQuery();
        //            }
        //        }
        //    }
        //}


        //private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        //{
        //    // If the part does not contain a SharedStringTable, create one.
        //    if (shareStringPart.SharedStringTable == null)
        //    {
        //        shareStringPart.SharedStringTable = new SharedStringTable();
        //    }
        //    int i = 0;
        //    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        //    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        //    {
        //        if (item.InnerText == text)
        //        {
        //            return i;
        //        }
        //        i++;
        //    }
        //    // The text does not exist in the part. Create the SharedStringItem and return its index.
        //    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        //    shareStringPart.SharedStringTable.Save();
        //    return i;
        //}
        //private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        //{
        //    // Add a new worksheet part to the workbook.
        //    WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //    newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        //    newWorksheetPart.Worksheet.Save();
        //    Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        //    string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
        //    // Get a unique ID for the new sheet.
        //    uint sheetId = 1;
        //    if (sheets.Elements<Sheet>().Count() > 0)
        //    {
        //        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        //    }
        //    string sheetName = "Sheet" + sheetId;
        //    // Append the new worksheet and associate it with the workbook.
        //    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        //    sheets.Append(sheet);
        //    workbookPart.Workbook.Save();
        //    return newWorksheetPart;
        //}


        //private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        //{
        //    Worksheet worksheet = worksheetPart.Worksheet;
        //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        //    string cellReference = columnName + rowIndex;
        //    // If the worksheet does not contain a row with the specified row index, insert one.
        //    Row row;
        //    if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        //    {
        //        row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        //    }
        //    else
        //    {
        //        row = new Row() { RowIndex = rowIndex };
        //        sheetData.Append(row);
        //    }
        //    // If there is not a cell with the specified column name, insert one.  
        //    if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        //    {
        //        return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        //    }
        //    else
        //    {
        //        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
        //        Cell refCell = null;
        //        foreach (Cell cell in row.Elements<Cell>())
        //        {
        //            if (cell.CellReference.Value.Length == cellReference.Length)
        //            {
        //                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
        //                {
        //                    refCell = cell;
        //                    break;
        //                }
        //            }
        //        }
        //        Cell newCell = new Cell() { CellReference = cellReference };
        //        row.InsertBefore(newCell, refCell);
        //        worksheet.Save();
        //        return newCell;
        //    }
        //}




        // method to read the password 
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