using SpreadsheetGear;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBO_RepeatedColumns_Finder
{
    class Program
    {
        static void Main(string[] args)
        {
          


            string folderPath = ConfigurationManager.AppSettings["FolderPath"].ToString();
            DirectoryInfo dirInfo = new DirectoryInfo(folderPath);

            DataTable errorDT = new DataTable();
            errorDT.Columns.Add("FileName", typeof(string));
            errorDT.Columns.Add("Sheet Name", typeof(string));
            errorDT.Columns.Add("RepeatedColumnNames", typeof(string));
            errorDT.Columns.Add("Description / Renamed To", typeof(string));

            //Error logging
            SpreadsheetGear.IWorkbook errorWorkbook = SpreadsheetGear.Factory.GetWorkbook();
            SpreadsheetGear.IWorksheet errorWorksheet = errorWorkbook.Worksheets.Add();
            SpreadsheetGear.IRange errorRange = errorWorksheet.Cells["A1"];
            errorWorksheet.Name = "Errors";


            bool nameProb = false;

            List<string> sheetNames = new List<string>();// used to concatinate all sheet names of a worksheet

            foreach (var item in dirInfo.GetFiles("*.xlsx"))
            {
                
               


                //IWorkbook workbook = GetWorkBook(item.FullName);
                IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbookSet().Workbooks.Open(item.FullName);
                workbook.WorkbookSet.Calculation = Calculation.Manual; // Setting to manual to avoid dirty read from excel formulae
                workbook.WorkbookSet.Calculate();

                Console.WriteLine(item);

                int count = 0;
              
                List<string> objSearchedCells = new List<string>();
            
                foreach (IWorksheet workSheet in workbook.Sheets)
                {
                    string[] categories = ConfigurationManager.AppSettings["sheetNames"].Split(',').Select(x => x.Trim().ToLower()).ToArray();

                    if (!categories.Contains(workSheet.Name.ToString().ToLower()))
                    {
                        foreach (var name in categories)
                        {
                            if(name.Contains(workSheet.Name.Split(' ')[0]))
                            {
                                errorDT.Rows.Add(item, workSheet.Name, string.Empty,"Sheet Name problem");
                                nameProb = true;
                            }
                        }
                        if (!nameProb)
                            errorDT.Rows.Add(item, workSheet.Name, string.Empty, "UnKnown Error");
                        continue;
                    }
                    else
                    {
                        sheetNames.Add(workSheet.Name);
                    }

                    if (workSheet.Name.ToLower().Trim().ToString().Equals("summary"))
                    {
                        IRange dateCell = null;
                        DateTime dateValue;
                        dateCell=workSheet.Cells.Find("For Month Ended:",null,FindLookIn.Values,LookAt.Part,SearchOrder.ByRows,SearchDirection.Next,false);

                        dateCell = dateCell.EndRight;
                        if (dateCell!=null)
                        {
                            double result;
                            if (double.TryParse(dateCell.Value.ToString(), out result))
                            {
                                dateValue = DateTime.FromOADate(Convert.ToDouble(result));

                                var today = DateTime.Today;
                                var month = new DateTime(today.Year, today.Month, 1);
                                var first = month.AddMonths(-1);
                                var last = month.AddDays(-1);

                                if (dateValue != last)
                                    errorDT.Rows.Add(item,workSheet.Name,string.Empty,"Date in the sheet is "+dateValue.ToShortDateString());

                            }

                        }
                        else
                        {
                            errorDT.Rows.Add(item, workSheet.Name, string.Empty, "Cannot Find Date");
                        }

                        continue;
                    }
                
                    
                    IRange startingCell = GetFirstCellRange(workSheet);
                    objSearchedCells.Clear();
                   
                  //Getting list of column names which are repeated in Sheet
                  var  objValLst = ((IEnumerable)startingCell.EntireRow.Value).Cast<string>().Where(x=>x!=null).ToList().GroupBy(w=>w).Where(g => g.Count() > 1).Select(g => g.Key).ToArray();


                    foreach (object cellValue in objValLst)
                    {
                        IRange searchedCell = null;
                   
                        objSearchedCells.Clear();
                        count = 0;
                        while (true)
                        {
                            searchedCell = workSheet.Cells.Find(cellValue.ToString(), searchedCell, SpreadsheetGear.FindLookIn.Values, SpreadsheetGear.LookAt.Part, SpreadsheetGear.SearchOrder.ByRows, SpreadsheetGear.SearchDirection.Next, false);
                            
                            if (Convert.ToInt32(startingCell.Address.ToString().Split('$')[2].ToString())< Convert.ToInt32(searchedCell.Address.ToString().Split('$')[2].ToString()))
                                break;

                       

                            if (searchedCell != null && (searchedCell.Address!=startingCell.Address))
                            {
                                if (searchedCell.Value.ToString().Equals(cellValue.ToString()))
                                {
                                 
                                    searchedCell.Value = cellValue.ToString() + "_" + count;
                                    count++;
                                    errorDT.Rows.Add(item,cellValue.ToString(),searchedCell.Value);
                                    break;
                                }
                                    
                            }


                            if (objSearchedCells.Contains(searchedCell.Address.ToString()))
                                break;
                            objSearchedCells.Add(searchedCell.Address.ToString());

                            
                        }
                    }
                    if(Convert.ToBoolean(ConfigurationManager.AppSettings["clearFormats"].ToString()))
                        workSheet.UsedRange.ClearFormats();

                }
                string fnlNames = null;
                foreach (var subSheetName in sheetNames)
                {
                    fnlNames = fnlNames + subSheetName + ",";
                }
                errorDT.Rows.Add(item, string.Empty, string.Empty,"Sheet has "+fnlNames.Trim(','));
                
                workbook.Save();
                workbook.Close();
                
            }

            string errorFolder = Path.Combine(ConfigurationManager.AppSettings["FolderPath"].ToString(), "ErrorFolder");
            string filePath = Path.Combine(errorFolder, "SBO_RepeatedColumn_List.xls");
          

            if (errorDT != null && errorDT.Rows.Count > 0)
            {

                if (!Directory.Exists(errorFolder))
                    Directory.CreateDirectory(errorFolder);

                errorRange.CopyFromDataTable(errorDT, SpreadsheetGear.Data.SetDataFlags.None);
                errorWorkbook.SaveAs(filePath, SpreadsheetGear.FileFormat.Excel8);
            }
            Console.WriteLine("Done");
            Console.ReadLine();
        }

        private static IWorkbook GetWorkBook(string filePath)
        {
            FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);

            return SpreadsheetGear.Factory.GetWorkbookSet()
                .Workbooks.OpenFromStream(fs);
        }

        private static void CreateDirectoryIfMissing(string errorFolder)
        {
            
           // filePath = Path.Combine(errorFolder,"SBO_RepeatedColumn_List.xls");

            if (!Directory.Exists(errorFolder))
                        Directory.CreateDirectory(errorFolder);
       
           // if (!File.Exists(filePath))
               //        File.Create(filePath).Dispose();    
            

        }
        public static IRange GetFirstCellRange(IWorksheet sheet)
        {
            IRange firstRow = null;
            foreach (IRange cell in sheet.Range["A1:A6"])
            {

                if (cell.Interior.Color == SpreadsheetGear.Color.FromArgb(-4210753))
                {
                    firstRow = cell.Range;
                }

            }
            return firstRow;
        }
    }
}
