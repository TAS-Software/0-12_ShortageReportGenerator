using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BomExcelGenerator
{
    public static class ImportManager
    {
        public static Dictionary<string, DateTime> ImportBFDInfo()
        {
            Dictionary<string, DateTime> import = new Dictionary<string, DateTime>();
            try
            {
                //string filepath = @"\\THAS-NAS01\RedirectedDocs$\chris.weeks\Desktop\MPStest.xlsm";
                string filepath = @"\\THAS-NAS01\DepartmentShares$\Production\Production Schedule\Master Schedule\MS Excel\MPS.xlsm";

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    WorkbookPart wbPart = doc.WorkbookPart;
                    int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count();
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
                    List<string> lols = doc.WorkbookPart.Workbook.Descendants<Sheet>().Select(x => x.Name.Value).ToList();
                    string relId = doc.WorkbookPart.Workbook.Descendants<Sheet>().First(s => s.Name.Value.Equals("TAS_Connect_Import")).Id;
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(relId)).Worksheet;
                    List<Row> rows = Worksheet.Descendants<Row>().Where(r => r.Hidden == null).ToList();

                    foreach (Row row in rows.Where(row => row != rows.First()).ToList())
                    {
                        try
                        {                          
                            string nameid = ((Cell)row.ChildElements.GetItem(0)).CellValue.InnerText;
                            int stringLocation = int.Parse(nameid);
                            string name = wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringLocation).InnerText;
                            DateTime date = DateTime.FromOADate(Convert.ToInt32(((Cell)row.ChildElements.GetItem(1)).CellValue.InnerText));
                            import.Add(name, date);
                        }
                        catch(Exception ex)
                        {
                            string exMsg = ex.Message;
                        }
                    }
                }
                return import;
            }
            catch (Exception Ex)
            {
                //SiteLogger.ReportError(Ex.Message);
                throw;
            }
        }    

        public static List<MilestonePriority> GetShortagesForBuilding()
        {
            List<MilestonePriority> milestones = new List<MilestonePriority>();
            try
            {
                string filepath = @"\\thas-report01\TASConnectImport\Shortages\milestones.xlsx";
                //string filepath = @"E:\\TASConnectImport\Shortages\milestones.xlsx";

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    WorkbookPart wbPart = doc.WorkbookPart;
                    int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count();
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
                    List<string> lols = doc.WorkbookPart.Workbook.Descendants<Sheet>().Select(x => x.Name.Value).ToList();
                    string relId = doc.WorkbookPart.Workbook.Descendants<Sheet>().First(s => s.Name.Value.Equals("Milestones")).Id;
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(relId)).Worksheet;
                    List<Row> rows = Worksheet.Descendants<Row>().Where(r => r.Hidden == null).ToList();

                    foreach (Row row in rows.ToList())
                    {
                        try
                        {
                            int pr = int.Parse(((Cell)row.ChildElements.GetItem(0)).CellValue.InnerText);                           

                            string msID = ((Cell)row.ChildElements.GetItem(1)).CellValue.InnerText;
                            int stringLocationMS = int.Parse(msID);
                            string ms = wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringLocationMS).InnerText;

                            // Only get a distinct list of all Milestones for shortage tracking.
                            if (!milestones.Select(x => x.Milestone).ToList().Contains(ms))
                                milestones.Add(new MilestonePriority { Priority = pr, Milestone = ms });
                        }
                        catch (Exception ex)
                        {
                            AppLogger.ReportError("Error encountered whilst processing milestone import row.  Details : " + ex.Message);
                            string exMsg = ex.Message;
                        }
                    }
                }
                return milestones;
            }
            catch (Exception ex)
            {
                AppLogger.ReportError("Error encountered whilst processing milestones import file.  Details : " + ex.Message);
                throw;
            }
        }


        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
    }
}