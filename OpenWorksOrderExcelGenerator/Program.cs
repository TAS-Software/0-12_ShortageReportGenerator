using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.Entity;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BomExcelGenerator
{
    class Program
    {
        public static DateTime targetDate = DateTime.Now.AddDays(1);
        public static string theDate = targetDate.ToString("yyyyMMdd");
        public static List<string> giants = new List<string>();
        static void Main(string[] args)
        {
            ProvideUserOptions();
        }

        private static void ProvideUserOptions()
        {
            PrintWelcomeMessage();
            ConsoleKeyInfo lol = Console.ReadKey();
            Console.WriteLine();
            int shortages = 0;

            if (lol.KeyChar == '1')
            {
                RunFullProcessing();
            }
            else if (lol.KeyChar == '2')
            {
                using (ReportDbEntities db = new ReportDbEntities())
                {
                    DateTime computeStart = DateTime.Now;
                    DateTime computeEnd = DateTime.Now;
                    Console.WriteLine();
                    Console.WriteLine("Starting Compute of BOM Part Totals at : " + DateTime.Now.ToString());
                    AppLogger.ReportInfo("Starting Compute of BOM Part Totals.");
                    computeStart = DateTime.Now;
                    db.Database.CommandTimeout = 72000;
                    int lmao = db.ConnectShortageTotalsBuilder();
                    computeEnd = DateTime.Now;
                    Console.WriteLine();
                    Console.WriteLine("Finished Compute of BOM Part Totals at : " + DateTime.Now.ToString());
                    AppLogger.ReportInfo("Finished Compute of BOM Part Totals.");
                    AppLogger.SendShortageReportGenerationUpdateEmail("All BOM Part Totals Per Milestone and Despatch Dates Are Now Calculated.", "BOM Part Totals Computation Complete.", "BOM Parts Totals Computation", "BOM Part Totals", "lol", 0, 0);
                }
                Console.ReadLine();
            }
            else if (lol.KeyChar == '3')
            {
                giants = new List<string> { "VT56-SS1", "VT59-ADHOC REVIEW1", "VT59-CDR SEATS", "VT59-CERT DYN T SEATS", "VT59-CERT STAT T SEATS", "VT59-DEV DYN T SEATS", "VT60-PDR SEATS", "VTMU2-SHOW SEATS", "VTMU-SALE SEATS2" };
                if (ContinueShortageReportsINDIVID(ref giants))
                {
                    AppLogger.ReportInfo("Finished Milestone Shortage Report Processing at: " + DateTime.Now.ToString());
                    AppLogger.SendShortageReportGenerationUpdateEmail("The Shortage Report Processing is now complete.", "Shortage Report Processing Complete", "Shortage Report Processing", "Shortage Report Processing", "processed", 0, 0);
                    Console.WriteLine();
                    Console.WriteLine("Finished Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                    //PerformMilestonesExports();
                    if (giants.Count() > 0)
                    {
                        ProcessShortageReportsGIANTS(giants);
                    }
                    PerfromSupplierBreakdownDumpAndExports();
                }
                else
                {
                    AppLogger.ReportError("The Shortage Processor Failed To Complete.");
                    Console.WriteLine(" ----- ");
                    Console.WriteLine("The Shortage Processor Failed To Complete.");
                    Console.WriteLine(" ----- ");
                }
            }
            else if (lol.KeyChar == '4')
            {
                //PerformMilestonesExports();
            }
            else if (lol.KeyChar == '5')
            {
                giants = new List<string>();
                if (ProcessShortageReportsINDIVID(ref giants))
                {
                    AppLogger.ReportInfo("Finished Milestone Shortage Report Processing at: " + DateTime.Now.ToString());
                    AppLogger.SendShortageReportGenerationUpdateEmail("The Shortage Report Processing is now complete.", "Shortage Report Processing Complete", "Shortage Report Processing", "Shortage Report Processing", "processed", 0, 0);
                    Console.WriteLine();
                    Console.WriteLine("Finished Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                    //PerformMilestonesExports();
                    if (giants.Count() > 0)
                    {
                        ProcessShortageReportsGIANTS(giants);
                    }
                }
                else
                {
                    AppLogger.ReportError("The Shortage Processor Failed To Complete.");
                    Console.WriteLine(" ----- ");
                    Console.WriteLine("The Shortage Processor Failed To Complete.");
                    Console.WriteLine(" ----- ");
                }
            }
            else if (lol.KeyChar == '6')
            {
                giants = new List<string> { "VT27-SS3", "PY01-SS7", "VT29-SS2", "VT31-SS2", "VT32-SS2", "VT23-SS4", "VT26-SS15", "VT33-SS4", "VT42-SS1", "VT42-SS1", "PY01-SS8", "VT10-SS28", "VT47-SS1", "VT27-SS4", "VT16-SS12", "VT49-SS12", "VT20-SS7", "VT38-SS7", "VT50-SS1", "VT16-SS12", "VT20-SS8" };
                if (giants.Count() > 0)
                {
                    ProcessShortageReportsGIANTS(giants);
                }
            }
            else if (lol.KeyChar == '7')
            {
                PerfromSupplierBreakdownDumpAndExports();
            }
            else if (lol.KeyChar == '0')
            {
                WaitForRunTime();
            }
            Console.ReadLine();
        }

        private static void WaitForRunTime()
        {
            TimeSpan _timeToRun = new TimeSpan(17, 30, 00);
            while (true)
            {
                TimeSpan timeNow = DateTime.Now.TimeOfDay;
                TimeSpan diff = _timeToRun - timeNow;
                Console.WriteLine(" ----- ");
                Console.WriteLine("Time Now is : " + timeNow.ToString());
                Console.WriteLine("Next Run Time is : " + _timeToRun.ToString());
                AppLogger.ReportInfo("Time Now is : " + timeNow.ToString());
                AppLogger.ReportInfo("Next Run Time is : " + _timeToRun.ToString());

                if (timeNow > _timeToRun) // run it the next day...
                {
                    var wait = new TimeSpan(1, 0, 0, 0) - diff.Duration();
                    Console.WriteLine("The Wait Window is : " + wait.ToString());
                    Console.WriteLine("We will sleep for this time.");
                    AppLogger.ReportInfo("The Wait Window is : " + wait.ToString());
                    AppLogger.ReportInfo("We will sleep for this time.");
                    Console.WriteLine(" ----- ");
                    System.Threading.Thread.Sleep(wait);
                }
                else // run it today
                {
                    Console.WriteLine("The Wait Window is : " + diff.Duration().ToString());
                    Console.WriteLine("We will sleep until this time.");
                    AppLogger.ReportInfo("The Wait Window is : " + diff.Duration().ToString());
                    AppLogger.ReportInfo("We will sleep for this time.");
                    Console.WriteLine(" ----- ");
                    System.Threading.Thread.Sleep(diff.Duration());
                }
                AppLogger.ReportInfo("We have reached the scheduled run time : " + DateTime.Now.TimeOfDay.ToString());
                RunFullProcessing();
            }
        }

        private static void RunFullProcessing()
        {
            if (RunShortageBuilder())
            {
                if (ProcessShortageReportsINDIVID(ref giants))
                {
                    AppLogger.ReportInfo("Finished Milestone Shortage Report Processing at: " + DateTime.Now.ToString());
                    AppLogger.SendShortageReportGenerationUpdateEmail("The Shortage Report Processing is now complete.", "Shortage Report Processing Complete", "Shortage Report Processing", "Shortage Report Processing", "processed", 0, 0);
                    Console.WriteLine();
                    Console.WriteLine("Finished Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                    //PerformMilestonesExports();
                    if (giants.Count() > 0)
                    {
                        ProcessShortageReportsGIANTS(giants);
                    }
                    PerfromSupplierBreakdownDumpAndExports();
                }
                else
                {
                    AppLogger.ReportError("The Shortage Processor Failed To Complete.");
                    Console.WriteLine(" ----- ");
                    Console.WriteLine("The Shortage Processor Failed To Complete.");
                    Console.WriteLine(" ----- ");
                }
            }
            else
            {
                AppLogger.ReportError("The Shortage Builder Failed To Run.");
                Console.WriteLine(" ----- ");
                Console.WriteLine("The Shortage Builder Failed To Run.");
                Console.WriteLine(" ----- ");
            }
        }

        private static List<SalesOrder> GetSOsFro123(double days)
        {
            try
            {
                using (thas01Entities mrpDB = new thas01Entities())
                {
                    DateTime targetDate = DateTime.Now.AddDays(days);
                    string res = string.Empty;
                    List<string> exclusions = new List<string> { "Spares Order", "SR", string.Empty, "", null };
                    List<SalesOrder> salesOrders = mrpDB.SalesOrders.Where(x => x.Status == 1 && !exclusions.Contains(x.SalesOrderTitle)
                    && !x.SalesOrderTitle.ToLower().Contains("spares") && x.SalesOrderDetails.Any(sd => sd.DespatchDate <= targetDate && sd.DespatchStatusID < 3)).ToList();
                    return salesOrders;
                }
            }
            catch (Exception ex)
            {
                AppLogger.ReportError("Error encountered whilst getting all SO's from 123 for the date range requested. Details : " + ex.Message);
                return new List<SalesOrder>();
            }
        }

        private static bool RunShortageBuilder()
        {
            try
            {
                List<string> milestones = GetScheduleOfSalesOrders();
                AppLogger.ReportInfo("Found " + milestones.Count + " milestones for import.");
                List<KeyValuePair<string, string>> failures = new List<KeyValuePair<string, string>>();
                List<string> successes = new List<string>();
                int numSuccesses = 0;
                DateTime start = DateTime.Now;
                DateTime end = DateTime.Now;
                DateTime computeStart = DateTime.Now;
                DateTime computeEnd = DateTime.Now;

                if (milestones != null && milestones.Count > 0)
                {
                    AppLogger.ReportInfo("Preparing for BOM Dump & Re-Compute.  Deleting all BOM Dump & BOM Part Total Records.", true);
                    using (ReportDbEntities db = new ReportDbEntities())
                    {
                        db.Database.ExecuteSqlCommand("truncate table BOMDump");
                        db.Database.ExecuteSqlCommand("truncate table BOMPartTotals");
                        db.Database.ExecuteSqlCommand("truncate table BOMWOCOPY");
                        db.Database.ExecuteSqlCommand("truncate table BOMPartIssues");
                        db.SaveChanges();
                    }

                    PerformBOMDump(milestones, successes, failures, numSuccesses);
                    end = DateTime.Now;
                    Console.WriteLine();
                    Console.WriteLine("Finished BOM Dump at : " + DateTime.Now.ToString());

                    using (ReportDbEntities db = new ReportDbEntities())
                    {
                        // If we have processed some BOM's we need to compute all the part totals
                        if (milestones.Count > 0)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Starting Compute of BOM Part Totals at : " + DateTime.Now.ToString());
                            AppLogger.ReportInfo("Starting Compute of BOM Part Totals.");
                            computeStart = DateTime.Now;
                            db.Database.CommandTimeout = 72000;
                            int lmao = db.ConnectShortageTotalsBuilder();
                            computeEnd = DateTime.Now;
                            Console.WriteLine();
                            Console.WriteLine("Finished Compute of BOM Part Totals at : " + DateTime.Now.ToString());
                            AppLogger.ReportInfo("Finished Compute of BOM Part Totals.");
                            AppLogger.SendShortageReportGenerationUpdateEmail("All BOM Part Totals Per Milestone and Despatch Dates Are Now Calculated.", "BOM Part Totals Computation Complete.", "BOM Parts Totals Computation", "BOM Part Totals", "lol", 0, 0);
                        }
                        else
                        {
                            AppLogger.ReportWarning("Found no milestones after BOM Dump Process");
                            return false;
                        }
                        // If we have processed some BOM's we need to COPY OVER THE WORKS ORDERS FOR THESE MILESTONES 
                        if (milestones.Count > 0)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Starting COPY of ALL WORKS ORDERS For Milestones in 12 Week Period at : " + DateTime.Now.ToString());
                            AppLogger.ReportInfo("Starting COPY of ALL WORKS ORDERS For Milestones in 12 Week Period ");
                            computeStart = DateTime.Now;
                            db.Database.CommandTimeout = 72000;
                            int lmao = db.ConnectShortageWOCopyWorksOrdersFromProd();
                            computeEnd = DateTime.Now;
                            Console.WriteLine();
                            Console.WriteLine("Finished COPY of ALL WORKS ORDERS For Milestones in 12 Week Period at : " + DateTime.Now.ToString());
                            AppLogger.ReportInfo("Finished COPY of ALL WORKS ORDERS For Milestones in 12 Week Period.");
                            //AppLogger.SendShortageReportGenerationUpdateEmail("All BOM Part Totals Per Milestone and Despatch Dates Are Now Calculated.", "BOM Part Totals Computation Complete.", "BOM Parts Totals Computation", "BOM Part Totals", "lol", 0, 0);
                        }
                        else
                        {
                            AppLogger.ReportWarning("Found no milestones after BOM Dump Process");
                            return false;
                        }
                        // If we have processed some BOM's we need to compute all of the Part Issues to OPEN WORKS ORDERS 
                        if (milestones.Count > 0)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Starting Compute of WO Issues at : " + DateTime.Now.ToString());
                            AppLogger.ReportInfo("Starting Compute of WO Issues.");
                            computeStart = DateTime.Now;
                            db.Database.CommandTimeout = 72000;
                            int lmao = db.ConnectShortageWOIssuedBuilder();
                            computeEnd = DateTime.Now;
                            Console.WriteLine();
                            Console.WriteLine("Finished Compute of WO Issues at : " + DateTime.Now.ToString());
                            AppLogger.ReportInfo("Finished Compute of WO Issues.");
                            //AppLogger.SendShortageReportGenerationUpdateEmail("All BOM Part Totals Per Milestone and Despatch Dates Are Now Calculated.", "BOM Part Totals Computation Complete.", "BOM Parts Totals Computation", "BOM Part Totals", "lol", 0, 0);
                        }
                        else
                        {
                            AppLogger.ReportWarning("Found no milestones after BOM Dump Process");
                            return false;
                        }
                        return true;
                    }
                }
                else
                {
                    AppLogger.ReportWarning("Found no milestones after BOM Dump Process");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(" ----- ");
                Console.WriteLine("Error encountered whilst generating BOM Dump and Parts Totals. Details : " + ex.Message);
                Console.WriteLine("Inner exception details : " + ex.InnerException.Message);
                Console.WriteLine(" ----- ");
                AppLogger.ReportError("Error encountered whilst generating BOM Dump and Parts Totals. Details : " + ex.Message);
                if (ex.InnerException != null)
                    AppLogger.ReportError("Inner Exception Details : " + ex.Message);
                return false;
            }
        }

        private static void PerformBOMDump(List<string> milestones, List<string> successes, List<KeyValuePair<string, string>> failures, int numSuccesses)
        {
            AppLogger.ReportInfo("Starting BOM Dump at : " + DateTime.Now.ToString());
            Console.WriteLine();
            Console.WriteLine("Starting BOM Dump at : " + DateTime.Now.ToString());
            using (thas01Entities mrpDB = new thas01Entities())
            {
                milestones.ForEach(m =>
                {
                    try
                    {
                        int lol = mrpDB.THAS_CONNECT_BOMCYCLE(m);
                        if (lol == 0)
                        {
                            AppLogger.ReportWarning("Could not find Milestone as Sales Order.  Milestone : " + m);
                            failures.Add(new KeyValuePair<string, string>(m, "Could not find Milestone as Sales Order"));
                        }
                        else
                        {
                            successes.Add(m);
                            numSuccesses++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error encountered whilst generating BOM Dump for " + m + ". Details : " + ex.Message);
                        if (ex.InnerException != null)
                            Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
                        AppLogger.ReportWarning("Error encountered whilst generating BOM Dump for " + m + ". Details : " + ex.Message);
                        failures.Add(new KeyValuePair<string, string>(m, ex.Message));
                    }
                });
            }
            Console.WriteLine();
            Console.WriteLine("Finished BOM Dump Process.");
            AppLogger.ReportInfo("Finished BOM Dump Process.");

            // Now we look at which milestones failed to process and transfer their BOM info.  Report this out on the BOM Transfer email.
            List<string> lmaoz = new List<string>();
            using (var lmao = new ReportDbEntities())
            {
                lmaoz = lmao.BOMDumps.Select(x => x.SalesOrderTitle).Distinct().ToList();
            }
            List<string> notFoundz = milestones.Except(lmaoz).ToList();
            Console.WriteLine("Milestones Without BOM Processed : " + notFoundz.Count.ToString());
            notFoundz.ForEach(x =>
            {
                AppLogger.ReportInfo("No BOM Processed for Milestone " + x);
            });
            string lollers = "<p>There were <b>" + milestones.Count + "</b> milestones within the reporting window. <b>" + lmaoz.Count + "</b> successful and <b>" + notFoundz.Count + "</b> failures.</p><p>The following milestones failed to have their BOM processed...</p> <ul><li>" + notFoundz.Aggregate((x, y) => x + "</li><li>" + y).ToString() + "</li></ul>";
            AppLogger.SendShortageReportGenerationUpdateEmail("The BOM Transfer from the 123 database into the TAS Connect database has completed successfully." + lollers, "BOM Transfer From 123 to TAS Connect Complete", "BOM Transfer", "Milestone", "transfers", lmaoz.Count, notFoundz.Count);
            GenerateSalesOrdersWithoutBOMsReport(notFoundz);
            return;
        }

        #region old_milestone_exporter

        //private static void PerformMilestonesExports()
        //{
        //    List<string> successes = new List<string>();
        //    List<string> failures = new List<string>();
        //    Console.WriteLine(); Console.WriteLine();
        //    Console.WriteLine("Starting Milestone Shortage Report Exports at : " + DateTime.Now.ToString());
        //    AppLogger.ReportInfo("Starting Milestone Shortage Report Exports at : " + DateTime.Now.ToString());
        //    ReportDbEntities db = new ReportDbEntities();
        //    db.Configuration.LazyLoadingEnabled = false;
        //    var milestoneNames = db.BOMDumps.Select(c => new { c.SalesOrderTitle, c.DespatchDate }).Distinct().OrderBy(x => x.DespatchDate).Select(x => x.SalesOrderTitle).ToList();
        //    Console.WriteLine();
        //    Console.WriteLine("Retrieved " + milestoneNames.Count + " Milestones to process Shortage Reports for.");
        //    AppLogger.ReportInfo("Retrieved " + milestoneNames.Count + " Milestones to process Shortage Reports for.");
        //    var milestoneNameslol = new List<string> { "VT41-SS2" };
        //    //string theDate = DateTime.Now.AddDays(1).ToString("yyyyMMdd");
        //    string theDate = DateTime.Now.ToString("yyyyMMdd");
        //    //theDate = new DateTime(2018, 5, 3).ToString("yyyyMMdd");
        //    foreach (var milestone in milestoneNameslol)
        //    {
        //        FileInfo fileInfo;
        //        if (CreateDirectoryStructure(milestone, out fileInfo, theDate))
        //        {
        //            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
        //            {
        //                var workSheet = excelPackage.Workbook.Worksheets.Add("FULLBOM");
        //                var resultSet = new List<ConnectShortageReport_Result>().ToList();
        //                try
        //                {
        //                    Console.WriteLine();
        //                    Console.WriteLine("Performing Shortage Report for : " + milestone);
        //                    AppLogger.ReportInfo("Performing Shortage Report for : " + milestone);

        //                    using (ReportDbEntities repDb = new ReportDbEntities())
        //                    {
        //                        repDb.Database.CommandTimeout = 1500;                              
        //                        resultSet = repDb.ConnectShortageReport(milestone).ToList();                                   
        //                        //string command = "exec ConnectShortageReport @so='" + milestone + "'";
        //                        //var lolz = repDb.Database.SqlQuery<ConnectShortageReport_Result>(command); 
        //                        //var unreal = lolz.ToList();
        //                    }
        //                    AppLogger.ReportInfo("Shortage Report for : " + milestone + " successfully retrieved.  Now generating export file.");
        //                    Console.WriteLine();
        //                    Console.WriteLine("Shortage Report for : " + milestone + " successfully retrieved.");
        //                    Console.WriteLine("Now generating export file.");
        //                }
        //                catch (Exception ex)
        //                {
        //                    Console.WriteLine();
        //                    Console.WriteLine("Exception encountered whilst generating shortage report for : " + milestone + ".");
        //                    Console.WriteLine("Exception Details : " + ex.Message + ".");
        //                    AppLogger.ReportError("Exception encountered whilst generating shortage report for : " + milestone + ".");
        //                    AppLogger.ReportError("Exception Details : " + ex.Message + ".");
        //                    if (ex.InnerException != null)
        //                    {
        //                        Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
        //                        AppLogger.ReportError("Inner Exception Details : " + ex.InnerException.Message + ".");
        //                    }
        //                    failures.Add(milestone);
        //                    continue;
        //                }
        //                try
        //                {
        //                    List<BOMExport> exports = new List<BOMExport>();
        //                    List<BOMLineShortage> shortages = new List<BOMLineShortage>();
        //                    decimal? priority = new decimal(0.0);
        //                    if (resultSet.Count() > 0)
        //                    {
        //                        resultSet.ForEach(check =>
        //                        {
        //                            priority = resultSet.Where(i => i.ComponentPart.Equals(check.ComponentPart) && i.DespatchDate == check.DespatchDate).Sum(val => val.Quantity);
        //                            var export = new BOMExport();
        //                            export.SalesOrderTitle = check.SalesOrderTitle;
        //                            export.DespatchDate = check.DespatchDate;
        //                            export.CustReqDate = check.CusReqDate;
        //                            export.MainPart = check.MainPart;
        //                            export.MainPartDescription = check.MainPartDescription;
        //                            export.ComponentPart = check.ComponentPart;
        //                            export.ComponentPartDesc = check.ComponentPartDescription;
        //                            export.ComponentMethod = check.ComponentMethod;
        //                            export.Responsibility = check.Responsibility;
        //                            export.ProductGroup = check.ProductGroup;
        //                            export.ResourceType = check.ResourceType;
        //                            export.ResourceCode = check.ResourceCode;
        //                            export.ResourceGroupName = check.ResourceGroupName;
        //                            export.UnitOfMeasure = check.UnitOfMeasure;
        //                            export.Quantity = check.Quantity.GetValueOrDefault();
        //                            export.TotalBOMQuantity = priority.GetValueOrDefault();
        //                            export.PriorDemand = check.RunningTotal.GetValueOrDefault();
        //                            export.TotalDemand = check.TotalPartDemand.GetValueOrDefault();
        //                            export.Stock = check.Stock.GetValueOrDefault();
        //                            export.WoQuantity = check.WOTotalQuantity.GetValueOrDefault();
        //                            export.WoOnTime = check.WOMeetDemand.GetValueOrDefault();
        //                            export.WoArriving = check.WOEarliest.HasValue ? check.WOEarliest.GetValueOrDefault() : DateTime.MinValue;
        //                            export.WoDelayInDays = check.WOEarliest.HasValue && check.DespatchDate.HasValue ? check.WOEarliest.GetValueOrDefault().Subtract(check.DespatchDate.GetValueOrDefault()).TotalDays.ToString() : "-";
        //                            export.PoQuantity = check.POTotalQuantity.GetValueOrDefault();
        //                            export.PoOnTime = check.POMeetDemand.GetValueOrDefault();
        //                            export.PoArriving = check.POEarliest.HasValue ? check.POEarliest.GetValueOrDefault() : DateTime.MinValue;
        //                            export.PoDelayInDays = check.POEarliest.HasValue && check.DespatchDate.HasValue ? check.POEarliest.GetValueOrDefault().Subtract(check.DespatchDate.GetValueOrDefault()).TotalDays.ToString() : "-";
        //                            export.Shortage = (check.Stock.GetValueOrDefault() + check.WOTotalQuantity.GetValueOrDefault() + check.POTotalQuantity.GetValueOrDefault()) - (priority.GetValueOrDefault() + check.RunningTotal.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
        //                            export.OnTime = (check.Stock.GetValueOrDefault() + check.WOMeetDemand.GetValueOrDefault() + check.POMeetDemand.GetValueOrDefault()) - (priority.GetValueOrDefault() + check.RunningTotal.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
        //                            exports.Add(export);
        //                            //shortages.Add(GetShortageForDB(export));
        //                        });

        //                        // Lets Save the Shortages for this milestone to the database for ease of querying...

        //                        // Now lets process the Excel export for the milestone...
        //                        var countz = 2;
        //                        foreach (var shortage in exports)
        //                        {
        //                            var shortageValue = shortage.Shortage;
        //                            decimal dShortValue = Convert.ToDecimal(shortageValue);
        //                            var onTimeValue = shortage.OnTime;
        //                            decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);
        //                            var woArriving = shortage.WoArriving;
        //                            var poArriving = shortage.PoArriving;

        //                            if (woArriving != DateTime.MinValue)
        //                            {
        //                                workSheet.Cells["V" + countz].Style.Numberformat.Format = "dd/MM/yyyy";
        //                            }
        //                            else
        //                            {
        //                                workSheet.Cells["V" + countz].Style.Numberformat.Format = "@";
        //                            }
        //                            if (poArriving != DateTime.MinValue)
        //                            {
        //                                workSheet.Cells["Z" + countz].Style.Numberformat.Format = "dd/MM/yyyy";
        //                            }
        //                            else
        //                            {
        //                                workSheet.Cells["Z" + countz].Style.Numberformat.Format = "@";
        //                            }
        //                            if (dShortValue >= 0)
        //                            {
        //                                workSheet.Cells["AB" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                                workSheet.Cells["AB" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
        //                                workSheet.Cells["AB" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
        //                            }
        //                            else
        //                            {
        //                                workSheet.Cells["AB" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                                workSheet.Cells["AB" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
        //                                workSheet.Cells["AB" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                            }
        //                            if (dOnTimeValue >= 0)
        //                            {
        //                                workSheet.Cells["AC" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                                workSheet.Cells["AC" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
        //                                workSheet.Cells["AC" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
        //                            }
        //                            else
        //                            {
        //                                workSheet.Cells["AC" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                                workSheet.Cells["AC" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
        //                                workSheet.Cells["AC" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                            }
        //                            countz++;
        //                        }

        //                        workSheet.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
        //                        workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
        //                        int rowCount = workSheet.Dimension.Rows;
        //                        string dateTimeFormat = "dd/MM/yyyy";
        //                        workSheet.Cells["A1:AC1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        workSheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
        //                        workSheet.Cells["A1:N1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                        workSheet.Cells["O1:R1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
        //                        workSheet.Cells["O1:R1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                        workSheet.Cells["S1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
        //                        workSheet.Cells["S1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                        workSheet.Cells["T1:W1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
        //                        workSheet.Cells["T1:W1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
        //                        workSheet.Cells["X1:AA1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
        //                        workSheet.Cells["X1:AA1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
        //                        workSheet.Cells["AB1:AC1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
        //                        foreach (var cell in workSheet.Cells["V2:V" + rowCount])
        //                        {
        //                            if (cell.Value.ToString() == "01/01/0001 00:00:00")
        //                            {
        //                                cell.Value = "-";
        //                            }
        //                        }
        //                        foreach (var cell in workSheet.Cells["Z2:Z" + rowCount])
        //                        {
        //                            if (cell.Value.ToString() == "01/01/0001 00:00:00")
        //                            {
        //                                cell.Value = "-";
        //                            }
        //                        }
        //                        workSheet.Cells["AB1:AC1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                        workSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
        //                        workSheet.Cells["C2:C" + rowCount].Style.Numberformat.Format = dateTimeFormat;

        //                        excelPackage.Save();
        //                        AppLogger.ReportInfo("Excel export for " + milestone + " successfully generated.");
        //                        Console.WriteLine();
        //                        Console.WriteLine("Excel export for " + milestone + " successfully generated.");
        //                        successes.Add(milestone);
        //                    }
        //                }
        //                catch (Exception exportEx)
        //                {
        //                    AppLogger.ReportError("Excel export for " + milestone + " has failed. Details: " + exportEx.Message);
        //                    Console.WriteLine();
        //                    Console.WriteLine("Excel export for " + milestone + " has failed. Details: " + exportEx.Message);
        //                    continue;
        //                }
        //            }
        //        }
        //    }
        //    AppLogger.ReportInfo("Finished Milestone Shortage Report Exports.");
        //    //AppLogger.SendShortageReportGenerationUpdateEmail("The Shortage Report Exports are now complete.", "Shortage Report Exports Complete", "Shortage Report Exports", "Shortage Report", "exports", milestoneNames.Count, failures.Count);
        //    Console.WriteLine();
        //    Console.WriteLine("Finished Milestone Shortage Report Exports at : " + DateTime.Now.ToString());
        //}
        #endregion

        private static Boolean ProcessShortageReports(ref int shortages)
        {
            try
            {
                Console.WriteLine("Starting Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                AppLogger.ReportInfo("Starting Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                using (ReportDbEntities repDb = new ReportDbEntities())
                {
                    repDb.Database.CommandTimeout = 72000;

                    Console.WriteLine("Truncating current BOMLineShortageYest table.");
                    AppLogger.ReportInfo("Truncating current BOMLineShortageYest table.");
                    repDb.Database.ExecuteSqlCommand("truncate table BOMLineShortageYest"); // remove all of yesterdays data
                    Console.WriteLine("Truncate of current BOMLineShortageYest table complete.");
                    AppLogger.ReportInfo("Truncate of current BOMLineShortageYest table complete.");

                    Console.WriteLine("Copying LineShortage -> Yest.");
                    AppLogger.ReportInfo("Copying LineShortage -> Yest.");
                    repDb.Database.ExecuteSqlCommand("insert into BOMLineShortageYest select * from BOMLineShortage"); // move todays info into yesterdays collection
                    Console.WriteLine("Copy of LineShortage -> Yest complete.");
                    AppLogger.ReportInfo("Copy of LineShortage -> Yest complete.");

                    Console.WriteLine("Truncating current BOMLineShortage table.");
                    AppLogger.ReportInfo("Truncating current BOMLineShortage table.");
                    repDb.Database.ExecuteSqlCommand("truncate table BOMLineShortage"); // clear todays data for fresh copy
                    Console.WriteLine("Truncate of current BOMLineShortage table complete.");
                    AppLogger.ReportInfo("Truncate of current BOMLineShortage table complete.");

                    shortages = repDb.ConnectShortageReportFULL();
                    if (shortages != 0)
                        return true;
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                Console.WriteLine("Exception Details : " + ex.Message + ".");
                AppLogger.ReportError("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                AppLogger.ReportError("Exception Details : " + ex.Message + ".");
                if (ex.InnerException != null)
                {
                    Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
                    AppLogger.ReportError("Inner Exception Details : " + ex.InnerException.Message + ".");
                }
                return false;
            }
        }

        private static Boolean ProcessShortageReportsINDIVID(ref List<string> giants)
        {
            try
            {
                List<string> successes = new List<string>();
                List<string> failures = new List<string>();
                List<string> milestoneNames = new List<string>();

                using (ReportDbEntities repDb = new ReportDbEntities())
                {
                    Console.WriteLine("Starting Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                    AppLogger.ReportInfo("Starting Milestone Shortage Report Processing at : " + DateTime.Now.ToString());

                    repDb.Database.CommandTimeout = 72000;

                    Console.WriteLine("Truncating current BOMLineShortageYest table.");
                    AppLogger.ReportInfo("Truncating current BOMLineShortageYest table.");
                    repDb.Database.ExecuteSqlCommand("truncate table BOMLineShortageYest"); // remove all of yesterdays data
                    Console.WriteLine("Truncate of current BOMLineShortageYest table complete.");
                    AppLogger.ReportInfo("Truncate of current BOMLineShortageYest table complete.");

                    Console.WriteLine("Copying LineShortage -> Yest.");
                    AppLogger.ReportInfo("Copying LineShortage -> Yest.");
                    repDb.Database.ExecuteSqlCommand("insert into BOMLineShortageYest select * from BOMLineShortage"); // move todays info into yesterdays collection
                    Console.WriteLine("Copy of LineShortage -> Yest complete.");
                    AppLogger.ReportInfo("Copy of LineShortage -> Yest complete.");

                    Console.WriteLine("Truncating current BOMLineShortage table.");
                    AppLogger.ReportInfo("Truncating current BOMLineShortage table.");
                    repDb.Database.ExecuteSqlCommand("truncate table BOMLineShortage"); // clear todays data for fresh copy
                    Console.WriteLine("Truncate of current BOMLineShortage table complete.");
                    AppLogger.ReportInfo("Truncate of current BOMLineShortage table complete.");

                    milestoneNames = repDb.BOMDumps.Select(c => new { c.SalesOrderTitle, c.DespatchDate }).Distinct().OrderBy(x => x.DespatchDate).Select(x => x.SalesOrderTitle).Distinct().ToList();
                    Console.WriteLine("Milestone Names Found : " + milestoneNames.Count());
                    Console.WriteLine("Milestone Names Found (Distinct) : " + milestoneNames.Distinct().Count());
                    Console.WriteLine();
                    Console.WriteLine("Retrieved " + milestoneNames.Count + " Milestones to process Shortage Reports for.");
                    AppLogger.ReportInfo("Retrieved " + milestoneNames.Count + " Milestones to process Shortage Reports for.");
                }

                List<BOMDump> dump = new List<BOMDump>();
                foreach (var milestone in milestoneNames.Distinct())
                {
                    using (ReportDbEntities repDb = new ReportDbEntities())
                    {
                        repDb.Configuration.AutoDetectChangesEnabled = false;
                        repDb.Configuration.ValidateOnSaveEnabled = false;
                        Console.WriteLine();
                        Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Beginning Processing of Shortages for : " + milestone);
                        AppLogger.ReportInfo(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Beginning Processing of Shortages for : " + milestone);
                        dump = repDb.BOMDumps.Where(x => x.SalesOrderTitle == milestone).ToList();
                        if (dump.Count() < 30000)
                        {
                            int shortages = repDb.ConnectShortageReportINDIVID(milestone);
                            if (shortages != 0)
                            {
                                Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Finished Processing Shortages for : " + milestone);
                                AppLogger.ReportInfo(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Finished Processing Shortages for : " + milestone);
                                successes.Add(milestone);
                            }
                            else
                            {
                                Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " ERROR Processing Shortages for : " + milestone);
                                AppLogger.ReportError(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " ERROR Processing Shortages for : " + milestone);
                                failures.Add(milestone);
                            }
                        }
                        else
                        {
                            Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Milestone : " + milestone + " is very large and will be saved for processing later.");
                            AppLogger.ReportError(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Milestone : " + milestone + " is very large and will be saved for processing later.");
                            giants.Add(milestone);
                        }
                        repDb.Database.Connection.Close();
                    }
                }
                Console.WriteLine("There was a total of " + giants.Count() + " large milestones which were not processed.");
                AppLogger.ReportInfo("There was a total of " + giants.Count() + " large milestones which were not processed.");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                Console.WriteLine("Exception Details : " + ex.Message + ".");
                AppLogger.ReportError("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                AppLogger.ReportError("Exception Details : " + ex.Message + ".");
                if (ex.InnerException != null)
                {
                    Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
                    AppLogger.ReportError("Inner Exception Details : " + ex.InnerException.Message + ".");
                }
                return false;
            }
        }

        private static Boolean ContinueShortageReportsINDIVID(ref List<string> giants)
        {
            try
            {
                List<string> successes = new List<string>();
                List<string> failures = new List<string>();
                List<string> milestoneNames = new List<string>();

                using (ReportDbEntities repDb = new ReportDbEntities())
                {
                    Console.WriteLine("*** Continuing Milestone Shortage Report Processing at : " + DateTime.Now.ToString());
                    AppLogger.ReportInfo("*** Continuing Milestone Shortage Report Processing at : " + DateTime.Now.ToString());

                    repDb.Database.CommandTimeout = 72000;
                    milestoneNames = giants;
                    Console.WriteLine("Milestone Names Found : " + milestoneNames.Count());
                    Console.WriteLine("Milestone Names Found (Distinct) : " + milestoneNames.Distinct().Count());
                    Console.WriteLine();
                    Console.WriteLine("Retrieved " + milestoneNames.Count + " Milestones to process Shortage Reports for.");
                    AppLogger.ReportInfo("Retrieved " + milestoneNames.Count + " Milestones to process Shortage Reports for.");
                }

                List<BOMDump> dump = new List<BOMDump>();
                foreach (var milestone in milestoneNames.Distinct())
                {
                    using (ReportDbEntities repDb = new ReportDbEntities())
                    {
                        repDb.Configuration.AutoDetectChangesEnabled = false;
                        repDb.Configuration.ValidateOnSaveEnabled = false;
                        Console.WriteLine();
                        Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Beginning Processing of Shortages for : " + milestone);
                        AppLogger.ReportInfo(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Beginning Processing of Shortages for : " + milestone);
                        dump = repDb.BOMDumps.Where(x => x.SalesOrderTitle == milestone).ToList();
                        int shortages = repDb.ConnectShortageReportINDIVID(milestone);
                        if (shortages != 0)
                        {
                            Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Finished Processing Shortages for : " + milestone);
                            AppLogger.ReportInfo(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Finished Processing Shortages for : " + milestone);
                            successes.Add(milestone);
                        }
                        else
                        {
                            Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " ERROR Processing Shortages for : " + milestone);
                            AppLogger.ReportError(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " ERROR Processing Shortages for : " + milestone);
                            failures.Add(milestone);
                        }
                        repDb.Database.Connection.Close();
                    }
                }
                Console.WriteLine("There was a total of " + giants.Count() + " large milestones which were not processed.");
                AppLogger.ReportInfo("There was a total of " + giants.Count() + " large milestones which were not processed.");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                Console.WriteLine("Exception Details : " + ex.Message + ".");
                AppLogger.ReportError("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                AppLogger.ReportError("Exception Details : " + ex.Message + ".");
                if (ex.InnerException != null)
                {
                    Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
                    AppLogger.ReportError("Inner Exception Details : " + ex.InnerException.Message + ".");
                }
                return false;
            }
        }

        private static Boolean ProcessShortageReportsGIANTS(List<string> giants)
        {
            try
            {
                List<string> successes = new List<string>();
                List<string> failures = new List<string>();

                List<BOMDump> dump = new List<BOMDump>();

                Console.WriteLine("There is a total of " + giants.Count() + " large milestones which will now be processed.");
                AppLogger.ReportInfo("There is a total of " + giants.Count() + " large milestones which will now be processed.");

                foreach (var milestone in giants)
                {
                    using (ReportDbEntities repDb = new ReportDbEntities())
                    {
                        Console.WriteLine();
                        Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Beginning Processing of Shortages for : " + milestone);
                        AppLogger.ReportInfo(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Beginning Processing of Shortages for : " + milestone);
                        dump = repDb.BOMDumps.Where(x => x.SalesOrderTitle == milestone).ToList();
                        if (dump.Count() >= 20000)
                        {
                            int shortages = repDb.ConnectShortageReportINDIVID(milestone);
                            if (shortages != 0)
                            {
                                Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Finished Processing Shortages for : " + milestone);
                                AppLogger.ReportInfo(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Finished Processing Shortages for : " + milestone);
                                successes.Add(milestone);
                            }
                            else
                            {
                                Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " ERROR Processing Shortages for : " + milestone);
                                AppLogger.ReportError(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " ERROR Processing Shortages for : " + milestone);
                                failures.Add(milestone);
                            }
                        }
                        else
                        {
                            //Console.WriteLine(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Milestone : " + milestone + " is very large and will saved for processing later.");
                            //AppLogger.ReportError(DateTime.Now.ToString("ddMMyy hh:mm:ss") + " Milestone : " + milestone + " is very large and will saved for processing later.");                            
                        }
                    }
                }
                //PerformMilestonesExports(giants);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                Console.WriteLine("Exception Details : " + ex.Message + ".");
                AppLogger.ReportError("Exception encountered whilst processing shortage reports into BOMLineShortage db Table");
                AppLogger.ReportError("Exception Details : " + ex.Message + ".");
                if (ex.InnerException != null)
                {
                    Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
                    AppLogger.ReportError("Inner Exception Details : " + ex.InnerException.Message + ".");
                }
                return false;
            }
        }

        private static void PerformMilestonesExports(List<string> miles = null)
        {
            List<string> successes = new List<string>();
            List<string> failures = new List<string>();
            Console.WriteLine(); Console.WriteLine();
            Console.WriteLine("Starting Milestone Shortage Report Exports at : " + DateTime.Now.ToString());
            AppLogger.ReportInfo("Starting Milestone Shortage Report Exports at : " + DateTime.Now.ToString());
            ReportDbEntities db = new ReportDbEntities();
            db.Configuration.LazyLoadingEnabled = false;
            var milestoneNames = new List<string>();
            if (miles != null && miles.Count() > 0)
                milestoneNames = miles;
            else
            {
                //milestoneNames = db.BOMDumps.Select(c => new { c.SalesOrderTitle, c.DespatchDate }).Distinct().OrderBy(x => x.DespatchDate).Select(x => x.SalesOrderTitle).Distinct().ToList();
                milestoneNames = db.BOMDumps.Where(x => x.DespatchDate <= new DateTime(2019, 1, 15)).Select(c => new { c.SalesOrderTitle, c.DespatchDate }).Distinct().OrderBy(x => x.DespatchDate).Select(x => x.SalesOrderTitle).Distinct().ToList();
                // OK so we have subtracted 28 days from the despatch date to show a PODD date.  But we have now run this for 0 - 12 weeks.  So each despatch date is -28 and we only 
                // want to show the 4 - 12 week exports so we add 56 days to this because 0 week reports are 28 days back so 4 weeks is 56 days back.
                //DateTime stepbackdate = DateTime.Now;
                //milestoneNames = db.BOMDumps.Where(bd => bd.DespatchDate >= stepbackdate).Select(c => new { c.SalesOrderTitle, c.DespatchDate }).Distinct().OrderBy(x => x.DespatchDate).Select(x => x.SalesOrderTitle).Distinct().ToList();
            }
            Console.WriteLine();
            Console.WriteLine("Retrieved " + milestoneNames.Count + " Milestones to export Shortage Reports for.");
            AppLogger.ReportInfo("Retrieved " + milestoneNames.Count + " Milestones to export Shortage Reports for.");

            //string theDate = DateTime.Now.ToString("yyyyMMdd");
            foreach (var milestone in milestoneNames.Distinct())
            {
                FileInfo fileInfo;
                if (CreateDirectoryStructure(milestone, out fileInfo, theDate))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                    {
                        var workSheet = excelPackage.Workbook.Worksheets.Add("FULLBOM");
                        var resultSet = new List<BOMLineShortage>().ToList();
                        try
                        {
                            Console.WriteLine();
                            Console.WriteLine("Performing Shortage Report for : " + milestone);
                            AppLogger.ReportInfo("Performing Shortage Report for : " + milestone);

                            using (ReportDbEntities repDb = new ReportDbEntities())
                            {
                                repDb.Database.CommandTimeout = 72000;
                                resultSet = repDb.BOMLineShortages.Where(x => x.SalesOrderTitle == milestone).ToList();
                            }
                            AppLogger.ReportInfo("Shortage Report for : " + milestone + " successfully retrieved.  Now generating export file.");
                            Console.WriteLine();
                            Console.WriteLine("Shortage Report for : " + milestone + " successfully retrieved.");
                            Console.WriteLine("Now generating export file.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Exception encountered whilst generating shortage report for : " + milestone + ".");
                            Console.WriteLine("Exception Details : " + ex.Message + ".");
                            AppLogger.ReportError("Exception encountered whilst generating shortage report for : " + milestone + ".");
                            AppLogger.ReportError("Exception Details : " + ex.Message + ".");
                            if (ex.InnerException != null)
                            {
                                Console.WriteLine("Inner Exception Details : " + ex.InnerException.Message + ".");
                                AppLogger.ReportError("Inner Exception Details : " + ex.InnerException.Message + ".");
                            }
                            failures.Add(milestone);
                            continue;
                        }
                        try
                        {
                            List<BOMExport> exports = new List<BOMExport>();
                            decimal? priority = new decimal(0.0);
                            if (resultSet.Count() > 0)
                            {
                                resultSet.ForEach(check =>
                                {
                                    priority = resultSet.Where(i => i.ComponentPart.Equals(check.ComponentPart) && i.DespatchDate == check.DespatchDate).Sum(val => val.Quantity);
                                    var export = new BOMExport();
                                    export.SalesOrderTitle = check.SalesOrderTitle;
                                    export.PODD = check.DespatchDate;
                                    export.CustReqDate = check.CustReqDate;
                                    export.MainPart = check.MainPart;
                                    export.MainPartDescription = check.MainPartDescription;
                                    export.ComponentPart = check.ComponentPart;
                                    export.ComponentPartDesc = check.ComponentPartDescription;
                                    export.ComponentMethod = check.ComponentMethod;
                                    export.Responsibility = check.Responsibility;
                                    export.ProductGroup = check.ProductGroup;
                                    export.ResourceType = check.ResourceType;
                                    export.ResourceCode = check.ResourceCode;
                                    export.ResourceGroupName = check.ResourceGroupName;
                                    export.UnitOfMeasure = check.UnitOfMeasure;
                                    export.Quantity = check.Quantity.GetValueOrDefault();
                                    export.TotalBOMQuantity = priority.GetValueOrDefault();
                                    export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                    export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                    export.Stock = check.Stock.GetValueOrDefault();
                                    export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                    export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                    export.WoArriving = check.WoArriving.HasValue ? check.WoArriving.GetValueOrDefault() : DateTime.MinValue;
                                    export.WoDelayInDays = check.WoArriving.HasValue && check.DespatchDate.HasValue ? check.WoArriving.GetValueOrDefault().Subtract(check.DespatchDate.GetValueOrDefault()).TotalDays.ToString() : "-";
                                    export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                    export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                    export.PoArriving = check.PoArriving.HasValue ? check.PoArriving.GetValueOrDefault() : DateTime.MinValue;
                                    export.PoDelayInDays = check.PoArriving.HasValue && check.DespatchDate.HasValue ? check.PoArriving.GetValueOrDefault().Subtract(check.DespatchDate.GetValueOrDefault()).TotalDays.ToString() : "-";
                                    //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (priority.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                    //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (priority.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                    decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (priority.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (priority.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                    decimal normalisedShortage = Shortage + WOIssued;
                                    decimal normalisedOnTime = OnTime + WOIssued;
                                    //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                    //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                    export.Shortage = normalisedShortage;
                                    export.OnTime = normalisedOnTime;
                                    export.WOIssued = WOIssued;
                                    exports.Add(export);
                                });

                                var countz = 2;
                                foreach (var shortage in exports)
                                {
                                    var shortageValue = shortage.Shortage;
                                    decimal dShortValue = Convert.ToDecimal(shortageValue);
                                    var onTimeValue = shortage.OnTime;
                                    decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);
                                    var woArriving = shortage.WoArriving;
                                    var poArriving = shortage.PoArriving;

                                    if (woArriving != DateTime.MinValue)
                                    {
                                        workSheet.Cells["V" + countz].Style.Numberformat.Format = "dd/MM/yyyy";
                                    }
                                    else
                                    {
                                        workSheet.Cells["V" + countz].Style.Numberformat.Format = "@";
                                    }
                                    if (poArriving != DateTime.MinValue)
                                    {
                                        workSheet.Cells["Z" + countz].Style.Numberformat.Format = "dd/MM/yyyy";
                                    }
                                    else
                                    {
                                        workSheet.Cells["Z" + countz].Style.Numberformat.Format = "@";
                                    }
                                    if (dShortValue >= 0)
                                    {
                                        workSheet.Cells["AB" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["AB" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        workSheet.Cells["AB" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        workSheet.Cells["AB" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["AB" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        workSheet.Cells["AB" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    if (dOnTimeValue >= 0)
                                    {
                                        workSheet.Cells["AC" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["AC" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        workSheet.Cells["AC" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        workSheet.Cells["AC" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["AC" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        workSheet.Cells["AC" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    countz++;
                                }

                                workSheet.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
                                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                                int rowCount = workSheet.Dimension.Rows;
                                string dateTimeFormat = "dd/MM/yyyy";
                                workSheet.Cells["A1:AD1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                                workSheet.Cells["A1:N1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["O1:R1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                workSheet.Cells["O1:R1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["S1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                workSheet.Cells["S1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["T1:W1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                                workSheet.Cells["T1:W1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                workSheet.Cells["X1:AA1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                                workSheet.Cells["X1:AA1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                workSheet.Cells["AB1:AC1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                workSheet.Cells["AD1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                foreach (var cell in workSheet.Cells["V2:V" + rowCount])
                                {
                                    if (cell.Value.ToString() == "01/01/0001 00:00:00")
                                    {
                                        cell.Value = "-";
                                    }
                                }
                                foreach (var cell in workSheet.Cells["Z2:Z" + rowCount])
                                {
                                    if (cell.Value.ToString() == "01/01/0001 00:00:00")
                                    {
                                        cell.Value = "-";
                                    }
                                }
                                workSheet.Cells["AB1:AD1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                workSheet.Cells["C2:C" + rowCount].Style.Numberformat.Format = dateTimeFormat;

                                excelPackage.Save();
                                AppLogger.ReportInfo("Excel export for " + milestone + " successfully generated.");
                                Console.WriteLine();
                                Console.WriteLine("Excel export for " + milestone + " successfully generated.");
                                successes.Add(milestone);
                            }
                        }
                        catch (Exception exportEx)
                        {
                            AppLogger.ReportError("Excel export for " + milestone + " has failed. Details: " + exportEx.Message);
                            Console.WriteLine();
                            Console.WriteLine("Excel export for " + milestone + " has failed. Details: " + exportEx.Message);
                            continue;
                        }
                    }
                }
            }
            AppLogger.ReportInfo("Finished Milestone Shortage Report Exports.");
            if (miles != null && miles.Count > 0)
                AppLogger.SendShortageReportGenerationUpdateEmail("The Shortage Report Exports of LARGE BOMs are now complete.", "Shortage Report Exports Complete (LARGE BOMs)", "Shortage Report Exports", "Shortage Report", "Large BOM exports", miles.Count, failures.Count);
            else
                AppLogger.SendShortageReportGenerationUpdateEmail("The Shortage Report Exports are now complete.", "Shortage Report Exports Complete", "Shortage Report Exports", "Shortage Report", "exports", milestoneNames.Count - giants.Count, failures.Count);

            Console.WriteLine();
            Console.WriteLine("Finished Milestone Shortage Report Exports at : " + DateTime.Now.ToString());
        }

        private static void PerfromSupplierBreakdownDumpAndExports()
        {
            AppLogger.ReportInfo("Starting Supplier Breakdown Reports at : " + DateTime.Now.ToString());
            Console.WriteLine();
            Console.WriteLine("Starting Supplier Breakdown Reports at : " + DateTime.Now.ToString());
            using (ReportDbEntities rDb = new ReportDbEntities())
            {
                // Perform data dumps for both Supplier and SupplierSO relations            
                //List<BOMShortagesSupplierBreakdown> supplier = rDb.BOMShortagesSupplierBreakdowns.ToList();
                //Console.WriteLine("Current Supplier Breakdown Rowcount : " + supplier.Count());
                //rDb.GetShortagesByPartDateSupplierINSERT();
                //supplier = rDb.BOMShortagesSupplierBreakdowns.ToList();
                //Console.WriteLine("New Supplier Breakdown Rowcount : " + supplier.Count());

                List<BOMShortagesSupplierBreakdownSO> supplierso = rDb.BOMShortagesSupplierBreakdownSOes.ToList();
                Console.WriteLine("Current Supplier Breakdown SO Rowcount : " + supplierso.Count());
                rDb.GetShortagesByPartDateSupplierSOINSERT();
                supplierso = rDb.BOMShortagesSupplierBreakdownSOes.ToList();
                Console.WriteLine("New Supplier Breakdown SO Rowcount : " + supplierso.Count());

                // Perform Export of Supplier and SupplierSO Reports
                Console.WriteLine("Starting Supplier Breakdown Export: " + DateTime.Now.ToString());
                SupplierBreakdown();
                SupplierBreakdownSO();
                GenerateSupplierBuyerReports();
                GenerateSupplierBuyerReportsSO();
                Console.WriteLine("Finished Supplier Breakdown Export : " + DateTime.Now.ToString());
            }

            AppLogger.SendShortageReportGenerationUpdateEmail("The Supplier Breakdown Report Processing has completed successfully.", "BOM Supplier Breakdown Reporting", "Supplier Breakdown", "Breakdown", "exports", 0, 0);
            return;
        }

        public static void SupplierBreakdown()
        {
            ReportDbEntities rdb = new ReportDbEntities();
            var owners = rdb.BOMShortageProductGroups.Include(x => x.BOMShortageOwners).ToList();
            //var supplierBuyers = rdb.BOMSupplierManageMatrices.ToList();
            var supplierBuyers = new List<BOMSupplierManageMatrix>();
            using (var connect = new ConnectDbEntities()) { supplierBuyers = connect.BOMSupplierManageMatrices.ToList(); }

            
            string theDateHours = DateTime.Now.ToString("yyyyMMdd HH.mm.ss");

            FileInfo fileInfo;
            if (CreateDirectoryStructure(out fileInfo, theDate, theDateHours, @"BOMShortagesSupplierBreakdown", "Shortage Reports", true))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    var workSheet = excelPackage.Workbook.Worksheets.Add("FULLBOM");
                    var resultSet = new List<BOMShortagesSupplierBreakdown>().ToList();
                    try
                    {
                        rdb.Database.CommandTimeout = 72000;
                        resultSet = rdb.BOMShortagesSupplierBreakdowns.ToList();
                    }
                    catch (Exception ex)
                    {
                        //return View();               
                    }
                    try
                    {
                        List<BOMShortageSupplierBreakdownObj> exports = new List<BOMShortageSupplierBreakdownObj>();
                        decimal? priority = new decimal(0.0);
                        if (resultSet.Count() > 0)
                        {
                            resultSet.ForEach(check =>
                            {
                                //var podd = check.DespatchDate.Value.AddDays(-14);
                                check.ProductGroup = check.ProductGroup == null ? string.Empty : check.ProductGroup;
                                var own = owners.SingleOrDefault(x => x.Name.ToLower().Equals(check.ProductGroup.ToLower()));
                                var ownz = string.IsNullOrWhiteSpace(check.SupplierName) && own != null ? own.BOMShortageOwners.First().Name : check.ProductGroup;
                                List<BOMSupplierManageMatrix> myBuyer = supplierBuyers.Where(x => x.Supplier.Equals(check.SupplierName)).ToList();
                                List<string> myBuyerz = myBuyer != null && myBuyer.Count() > 0 ? myBuyer.Select(byz => byz.BuyerName).ToList() : new List<string> { string.Empty };
                                var export = new BOMShortageSupplierBreakdownObj();
                                //export.DespatchDate = check.DespatchDate.Value.AddDays(14);
                                export.DespatchDate = check.DespatchDate.Value.AddDays(42); // used to be 14 days back - the podd is now 28 days back in the SQL procedure. -- it is now 42 days.
                                export.PODD = check.DespatchDate;
                                export.ComponentPart = check.ComponentPart;
                                export.ComponentPartDescription = check.ComponentPartDescription;
                                export.ComponentMethod = check.ComponentMethod;
                                export.ProductGroup = check.ProductGroup;
                                export.Responsibility = check.Responsibility;
                                export.LeadTime = check.LeadTime.Value;
                                export.SupplierName = !string.IsNullOrWhiteSpace(check.SupplierName) ? check.SupplierName : ownz;
                                export.Quantity = check.Quantity.GetValueOrDefault();
                                export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                export.Stock = check.Stock.GetValueOrDefault();
                                export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                export.POArriving = check.POArriving.GetValueOrDefault();
                                export.PODelayInDays = check.PODelayInDays;
                                //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                //decimal normalisedShortage = Math.Abs(Shortage) - WOIssued;
                                //decimal normalisedOnTime = Math.Abs(OnTime) - WOIssued;
                                decimal normalisedShortage = Shortage + WOIssued;
                                decimal normalisedOnTime = OnTime + WOIssued;
                                //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                export.Shortage = normalisedShortage;
                                export.OnTime = normalisedOnTime;
                                export.UnitCost = check.UnitCost.Value;
                                export.BuyerName = myBuyerz.Aggregate((i, j) => i + ", " + j);
                                export.ReportDate = DateTime.Now.Date.ToShortDateString();
                                export.WOIssued = WOIssued;
                                exports.Add(export);
                            });

                            var countz = 2;
                            foreach (var shortage in exports)
                            {
                                var shortageValue = shortage.Shortage;
                                decimal dShortValue = Convert.ToDecimal(shortageValue);
                                var onTimeValue = shortage.OnTime;
                                decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);

                                if (dShortValue >= 0)
                                {
                                    workSheet.Cells["T" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells["T" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                    workSheet.Cells["T" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else
                                {
                                    workSheet.Cells["T" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells["T" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                    workSheet.Cells["T" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                }
                                if (dOnTimeValue >= 0)
                                {
                                    workSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                    workSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else
                                {
                                    workSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                    workSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                }
                                countz++;
                            }

                            int rowCount = workSheet.Dimension.Rows;
                            string dateTimeFormat = "dd/MM/yyyy";
                            workSheet.Cells["A1:Y1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells["A1:Y1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                            workSheet.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                            workSheet.Cells["A1:I1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            workSheet.Cells["J1:L1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                            workSheet.Cells["J1:L1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            workSheet.Cells["M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                            workSheet.Cells["M1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            workSheet.Cells["N1:O1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                            workSheet.Cells["N1:O1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                            workSheet.Cells["P1:S1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                            workSheet.Cells["P:S1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                            workSheet.Cells["T1:U1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                            workSheet.Cells["T1:U1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            workSheet.Cells["A2:A" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                            workSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                            workSheet.Cells["W2:W" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                            string excelName = "BOMShortagesSupplierBreakdown" + DateTime.Now.ToString("dd-MM-yy HH.mm.ss tt");
                            workSheet.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
                            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                            excelPackage.Save();

                            //using (var memoryStream = new MemoryStream())
                            //{
                            //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            //    Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
                            //    excelPackage.SaveAs(memoryStream);
                            //    memoryStream.WriteTo(Response.OutputStream);
                            //    Response.Flush();
                            //    Response.End();
                            //}
                        }
                        return;
                    }
                    catch (Exception exportEx)
                    {
                        string exmsg = exportEx.Message;
                        return;
                    }
                }
            }

            
        }

        public static void SupplierBreakdownSO()
        {
            //ReportDbEntities rdb = new ReportDbEntities();
            var supplierBuyers = new List<BOMSupplierManageMatrix>();
            using (var connect = new ConnectDbEntities()) { supplierBuyers = connect.BOMSupplierManageMatrices.ToList(); }

            using (ReportDbEntities rdb = new ReportDbEntities())
            {
                var owners = rdb.BOMShortageProductGroups.Include(x => x.BOMShortageOwners).ToList();

                string theDateHours = DateTime.Now.ToString("yyyyMMdd HH.mm.ss");

                FileInfo fileInfo;
                if (CreateDirectoryStructure(out fileInfo, theDate, theDateHours, @"BOMShortagesSupplierBreakdownBySO", "Shortage Reports", true))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                    {
                        var workSheet = excelPackage.Workbook.Worksheets.Add("FULLBOM");
                        var resultSet = new List<BOMShortagesSupplierBreakdownSO>().ToList();
                        try
                        {
                            rdb.Database.CommandTimeout = 72000;
                            resultSet = rdb.BOMShortagesSupplierBreakdownSOes.ToList();
                        }
                        catch (Exception ex)
                        {
                            //return View();               
                        }
                        try
                        {
                            List<BOMShortageSupplierBreakdownSOObj> exports = new List<BOMShortageSupplierBreakdownSOObj>();
                            decimal? priority = new decimal(0.0);
                            if (resultSet.Count() > 0)
                            {
                                resultSet.ForEach(check =>
                                {
                                    //var podd = check.DespatchDate.Value.AddDays(-14);
                                    check.ProductGroup = check.ProductGroup == null ? string.Empty : check.ProductGroup;
                                    var own = owners.SingleOrDefault(x => x.Name.ToLower().Equals(check.ProductGroup.ToLower()));
                                    var ownz = string.IsNullOrWhiteSpace(check.SupplierName) && own != null ? own.BOMShortageOwners.First().Name : check.ProductGroup;
                                    //var myBuyer = supplierBuyers.SingleOrDefault(x => x.Supplier.Equals(check.SupplierName));
                                    //var myBuyerz = myBuyer != null ? myBuyer.BuyerName : string.Empty;
                                    List<BOMSupplierManageMatrix> myBuyer = supplierBuyers.Where(x => x.Supplier.Equals(check.SupplierName)).ToList();
                                    List<string> myBuyerz = myBuyer != null && myBuyer.Count() > 0 ? myBuyer.Select(byz => byz.BuyerName).ToList() : new List<string> { string.Empty };
                                    var export = new BOMShortageSupplierBreakdownSOObj();
                                    export.SalesOrderTitle = check.SalesOrderTitle;
                                    //export.DespatchDate = check.DespatchDate.Value.AddDays(14);
                                    export.DespatchDate = check.DespatchDate.Value.AddDays(42);  // used to be 14 days back - the podd is now 28 days back in the SQL procedure. -- it is now 42 days.
                                    export.PODD = check.DespatchDate;
                                    export.ComponentPart = check.ComponentPart;
                                    export.ComponentPartDescription = check.ComponentPartDescription;
                                    export.ComponentMethod = check.ComponentMethod;
                                    export.ProductGroup = check.ProductGroup;
                                    export.Responsibility = check.Responsibility;
                                    export.LeadTime = check.LeadTime.Value;
                                    //export.SupplierName = !string.IsNullOrWhiteSpace(check.SupplierName) ? check.SupplierName : check.Responsibility;
                                    export.SupplierName = !string.IsNullOrWhiteSpace(check.SupplierName) ? check.SupplierName : ownz;
                                    export.Quantity = check.Quantity.GetValueOrDefault();
                                    export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                    export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                    export.Stock = check.Stock.GetValueOrDefault();
                                    export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                    export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                    export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                    export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                    export.POArriving = check.POArriving.GetValueOrDefault();
                                    export.PODelayInDays = check.PODelayInDays;
                                    //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                    //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                    decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal Shortage2 = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault());
                                    decimal OnTime2 = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault());
                                    decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                    //decimal normalisedShortage = Math.Abs(Shortage) - WOIssued;
                                    //decimal normalisedOnTime = Math.Abs(OnTime) - WOIssued;
                                    decimal normalisedShortage = Shortage + WOIssued;
                                    decimal normalisedOnTime = OnTime + WOIssued;
                                    //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                    //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                    export.Shortage = normalisedShortage;
                                    export.OnTime = normalisedOnTime;
                                    export.UnitCost = check.UnitCost.Value;
                                    //export.BuyerName = myBuyerz;
                                    export.BuyerName = myBuyerz.Aggregate((i, j) => i + ", " + j);
                                    export.ReportDate = DateTime.Now.Date.ToShortDateString();
                                    export.WOIssued = WOIssued;
                                    exports.Add(export);
                                });

                                var countz = 2;
                                int shortbad = exports.Where(x => x.Shortage == null).Count();
                                int delaybad = exports.Where(x => x.OnTime == null).Count();

                                foreach (var shortage in exports)
                                {
                                    var shortageValue = shortage.Shortage;
                                    decimal dShortValue = Convert.ToDecimal(shortageValue);
                                    var onTimeValue = shortage.OnTime;
                                    decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);

                                    if (dShortValue >= 0)
                                    {
                                        workSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        workSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        workSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        workSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    if (dOnTimeValue >= 0)
                                    {
                                        workSheet.Cells["V" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["V" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        workSheet.Cells["V" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        workSheet.Cells["V" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells["V" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        workSheet.Cells["V" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    countz++;
                                }

                                int rowCount = workSheet.Dimension.Rows;
                                string dateTimeFormat = "dd/MM/yyyy";
                                workSheet.Cells["A1:Z1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells["A1:Z1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DodgerBlue);
                                workSheet.Cells["A1:J1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                                workSheet.Cells["A1:J1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["K1:M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                workSheet.Cells["K1:M1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["N1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                workSheet.Cells["N1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["O1:P1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                                workSheet.Cells["O1:P1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                workSheet.Cells["Q1:T1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                                workSheet.Cells["Q:T1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                workSheet.Cells["U1:V1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                workSheet.Cells["U1:V1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                workSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                workSheet.Cells["C2:C" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                workSheet.Cells["X2:X" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                string excelName = "BOMShortagesSupplierBreakdownSO" + DateTime.Now.ToString("dd-MM-yy HH.mm.ss tt");
                                workSheet.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
                                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                                excelPackage.Save();

                                //Now persist this out to the database.
                                rdb.Database.ExecuteSqlCommand("truncate table BOMShortagesSupplierBreakdownSOOutputBI"); // clear todays data for fresh copy
                                rdb.SaveChanges();
                                List<BOMShortagesSupplierBreakdownSOOutputBI> outputs = new List<BOMShortagesSupplierBreakdownSOOutputBI>();
                                exports.ForEach(ex =>
                                {
                                    BOMShortagesSupplierBreakdownSOOutputBI output = new BOMShortagesSupplierBreakdownSOOutputBI();
                                    output.SalesOrderTitle = ex.SalesOrderTitle;
                                    output.DespatchDate = ex.DespatchDate;
                                    output.PODD = ex.PODD;
                                    output.ComponentPart = ex.ComponentPart;
                                    output.ComponentPartDescription = ex.ComponentPartDescription;
                                    output.ComponentMethod = ex.ComponentMethod;
                                    output.ProductGroup = ex.ProductGroup;
                                    output.Responsibility = ex.Responsibility;
                                    output.LeadTime = ex.LeadTime;
                                    output.SupplierName = ex.SupplierName;
                                    output.Quantity = ex.Quantity;
                                    output.PriorDemand = ex.PriorDemand;
                                    output.TotalDemand = ex.TotalDemand;
                                    output.Stock = ex.Stock;
                                    output.WoQuantity = ex.WoQuantity;
                                    output.WoOnTime = ex.WoOnTime;
                                    output.PoQuantity = ex.PoQuantity;
                                    output.PoOnTime = ex.PoOnTime;
                                    output.PoArriving = ex.POArriving;
                                    output.PoDelayInDays = ex.PODelayInDays;
                                    output.Shortage = ex.Shortage;
                                    output.OnTime = ex.OnTime;
                                    output.WOIssued = ex.WOIssued;
                                    output.ReportDate = DateTime.Now.Date;
                                    output.UnitCost = ex.UnitCost;
                                    output.BuyerName = ex.BuyerName;
                                    outputs.Add(output);
                                });
                                rdb.BOMShortagesSupplierBreakdownSOOutputBIs.AddRange(outputs);
                                rdb.SaveChanges();


                                //using (var memoryStream = new MemoryStream())
                                //{
                                //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                //    Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
                                //    excelPackage.SaveAs(memoryStream);
                                //    memoryStream.WriteTo(Response.OutputStream);
                                //    Response.Flush();
                                //    Response.End();
                                //}
                            }
                            return;
                        }
                        catch (Exception exportEx)
                        {
                            string exmsg = exportEx.Message;
                            return;
                        }
                    }
                }

                
            }
        }

        public static void GenerateSupplierBuyerReports()
        {
            using (var rdb = new ReportDbEntities())
            {
                //var owners = rdb.BOMShortageProductGroups.Include(x => x.BOMShortageOwners).ToList();
                //var supplierBuyers = rdb.BOMSupplierManageMatrices.ToList();
                var supplierBuyers = new List<BOMSupplierManageMatrix>();
                using (var connect = new ConnectDbEntities()) { supplierBuyers = connect.BOMSupplierManageMatrices.ToList(); }

                List<Buyer> buyerInfo = new List<Buyer>();

                string theDateHours = DateTime.Now.ToString("yyyyMMdd HH.mm.ss");

                FileInfo fileInfo;
                if (CreateDirectoryStructure(out fileInfo, theDate, theDateHours, @"ShortagesBuyerWorkToList", "Shortage Reports", true))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                    {
                        var workSheet = excelPackage.Workbook.Worksheets.Add("Buyers Overview");
                        var resultSet = new List<BOMShortagesSupplierBreakdown>().ToList();
                        try
                        {
                            rdb.Database.CommandTimeout = 72000;
                            resultSet = rdb.BOMShortagesSupplierBreakdowns.ToList();
                        }
                        catch (Exception ex)
                        {
                            //return View();               
                        }
                        try
                        {
                            supplierBuyers.GroupBy(x => x.BuyerName).ToList().ForEach(buyer =>
                            {
                                var buyerSheet = excelPackage.Workbook.Worksheets.Add(buyer.First().BuyerName);
                                var suppliers = buyer.Select(b => b.Supplier).ToList();
                                var supplierShorts = resultSet.Where(r => suppliers.Contains(r.SupplierName)).ToList();
                                buyerInfo.Add(new Buyer { BuyerName = buyer.First().BuyerName, SupplierCount = suppliers.Count(), ShortageCount = supplierShorts.Count() });

                                List<BOMShortageSupplierBreakdownObj> exports2 = new List<BOMShortageSupplierBreakdownObj>();
                                if (supplierShorts.Count() > 0)
                                {
                                    supplierShorts.ForEach(check =>
                                    {
                                        check.ProductGroup = check.ProductGroup == null ? string.Empty : check.ProductGroup;
                                        var export = new BOMShortageSupplierBreakdownObj();
                                        export.DespatchDate = check.DespatchDate.Value.AddDays(42);
                                        export.PODD = check.DespatchDate;
                                        export.ComponentPart = check.ComponentPart;
                                        export.ComponentPartDescription = check.ComponentPartDescription;
                                        export.ComponentMethod = check.ComponentMethod;
                                        export.ProductGroup = check.ProductGroup;
                                        export.Responsibility = check.Responsibility;
                                        export.LeadTime = check.LeadTime.Value;
                                        export.SupplierName = check.SupplierName;
                                        export.Quantity = check.Quantity.GetValueOrDefault();
                                        export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                        export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                        export.Stock = check.Stock.GetValueOrDefault();
                                        export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                        export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                        export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                        export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                        export.POArriving = check.POArriving.GetValueOrDefault();
                                        export.PODelayInDays = check.PODelayInDays;
                                        //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                        //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                        decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                        decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                        decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                        //decimal normalisedShortage = Math.Abs(Shortage) - WOIssued;
                                        //decimal normalisedOnTime = Math.Abs(OnTime) - WOIssued;
                                        decimal normalisedShortage = Shortage + WOIssued;
                                        decimal normalisedOnTime = OnTime + WOIssued;
                                        //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                        //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                        export.Shortage = normalisedShortage;
                                        export.OnTime = normalisedOnTime;
                                        export.WOIssued = WOIssued;
                                        export.UnitCost = check.UnitCost.Value;
                                        export.ReportDate = DateTime.Now.Date.ToShortDateString();
                                        export.BuyerName = buyer.First().BuyerName;
                                        if (export.OnTime < 0)
                                            exports2.Add(export);
                                    });

                                    var countz = 2;
                                    foreach (var shortage in exports2)
                                    {
                                        var shortageValue = shortage.Shortage;
                                        decimal dShortValue = Convert.ToDecimal(shortageValue);
                                        var onTimeValue = shortage.OnTime;
                                        decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);

                                        if (dShortValue >= 0)
                                        {
                                            buyerSheet.Cells["T" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["T" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                            buyerSheet.Cells["T" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                        }
                                        else
                                        {
                                            buyerSheet.Cells["T" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["T" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            buyerSheet.Cells["T" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                        }
                                        if (dOnTimeValue >= 0)
                                        {
                                            buyerSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                            buyerSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                        }
                                        else
                                        {
                                            buyerSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            buyerSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                        }
                                        if (shortage.PODD.Value.AddDays(-shortage.LeadTime) < DateTime.Now)
                                        {
                                            buyerSheet.Cells["H" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["H" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            buyerSheet.Cells["H" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                        }
                                        countz++;
                                    }

                                    int rowCount = exports2.Count();
                                    string dateTimeFormat = "dd/MM/yyyy";
                                    buyerSheet.Cells["A1:Y1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    buyerSheet.Cells["A1:Y1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                    buyerSheet.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                                    buyerSheet.Cells["A1:I1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["J1:L1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                    buyerSheet.Cells["J1:L1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                    buyerSheet.Cells["M1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["N1:O1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                                    buyerSheet.Cells["N1:O1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    buyerSheet.Cells["P1:S1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                                    buyerSheet.Cells["P:S1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    buyerSheet.Cells["T1:U1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                    buyerSheet.Cells["T1:U1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["A2:A" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["R2:R" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["W2:W" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["A1"].LoadFromCollection(exports2, true, OfficeOpenXml.Table.TableStyles.Medium2);
                                    buyerSheet.Cells[buyerSheet.Dimension.Address].AutoFitColumns();
                                };
                            });

                            // Group unmanaged suppliers together...
                            var unmanagedSheet = excelPackage.Workbook.Worksheets.Add("Unassigned Suppliers");
                            var supplierShorts2 = resultSet.Where(r => !supplierBuyers.Select(sup => sup.Supplier).ToList().Contains(r.SupplierName)).ToList();
                            buyerInfo.Add(new Buyer { BuyerName = "Unassigned Suppliers", SupplierCount = supplierShorts2.Select(um => um.SupplierName).ToList().Distinct().Count(), ShortageCount = supplierShorts2.Count() });
                            workSheet.Cells["A1"].LoadFromCollection(buyerInfo, true, OfficeOpenXml.Table.TableStyles.Medium2);
                            workSheet.Cells["A1:C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells["A1:C1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                            workSheet.Cells["A1"].Value = "Buyer Name";
                            workSheet.Cells["B1"].Value = "Supplier Count";
                            workSheet.Cells["C1"].Value = "Shortage Count";
                            workSheet.Column(1).Width = 20;
                            workSheet.Column(2).Width = 20;
                            workSheet.Column(3).Width = 20;

                            List<BOMShortageSupplierBreakdownObj> exports = new List<BOMShortageSupplierBreakdownObj>();
                            decimal? priority = new decimal(0.0);
                            if (supplierShorts2.Count() > 0)
                            {
                                supplierShorts2.ForEach(check =>
                                {
                                    check.ProductGroup = check.ProductGroup == null ? string.Empty : check.ProductGroup;
                                    var export = new BOMShortageSupplierBreakdownObj();
                                    export.DespatchDate = check.DespatchDate.Value.AddDays(42);
                                    export.PODD = check.DespatchDate;
                                    export.ComponentPart = check.ComponentPart;
                                    export.ComponentPartDescription = check.ComponentPartDescription;
                                    export.ComponentMethod = check.ComponentMethod;
                                    export.ProductGroup = check.ProductGroup;
                                    export.Responsibility = check.Responsibility;
                                    export.LeadTime = check.LeadTime.Value;
                                    export.SupplierName = check.SupplierName;
                                    export.Quantity = check.Quantity.GetValueOrDefault();
                                    export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                    export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                    export.Stock = check.Stock.GetValueOrDefault();
                                    export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                    export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                    export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                    export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                    export.POArriving = check.POArriving.GetValueOrDefault();
                                    export.PODelayInDays = check.PODelayInDays;
                                    //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                    //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                    decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                    //decimal normalisedShortage = Math.Abs(Shortage) - WOIssued;
                                    //decimal normalisedOnTime = Math.Abs(OnTime) - WOIssued;
                                    decimal normalisedShortage = Shortage + WOIssued;
                                    decimal normalisedOnTime = OnTime + WOIssued;
                                    //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                    //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                    export.Shortage = normalisedShortage;
                                    export.OnTime = normalisedOnTime;
                                    export.WOIssued = WOIssued;
                                    export.UnitCost = check.UnitCost.Value;
                                    export.ReportDate = DateTime.Now.Date.ToShortDateString();
                                    exports.Add(export);
                                });

                                var countz = 2;
                                foreach (var shortage in exports)
                                {
                                    var shortageValue = shortage.Shortage;
                                    decimal dShortValue = Convert.ToDecimal(shortageValue);
                                    var onTimeValue = shortage.OnTime;
                                    decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);

                                    if (dShortValue >= 0)
                                    {
                                        unmanagedSheet.Cells["T" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["T" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        unmanagedSheet.Cells["T" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        unmanagedSheet.Cells["T" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["T" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        unmanagedSheet.Cells["T" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    if (dOnTimeValue >= 0)
                                    {
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        unmanagedSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        unmanagedSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    if (shortage.PODD.Value.AddDays(-shortage.LeadTime) < DateTime.Now)
                                    {
                                        unmanagedSheet.Cells["H" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["H" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        unmanagedSheet.Cells["H" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    countz++;
                                }

                                int rowCount = exports.Count();
                                string dateTimeFormat = "dd/MM/yyyy";
                                unmanagedSheet.Cells["A1:Y1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                unmanagedSheet.Cells["A1:Y1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                unmanagedSheet.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                                unmanagedSheet.Cells["A1:I1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["J1:L1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                unmanagedSheet.Cells["J1:L1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                unmanagedSheet.Cells["M1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["N1:O1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                                unmanagedSheet.Cells["N1:O1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                unmanagedSheet.Cells["P1:S1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                                unmanagedSheet.Cells["P:S1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                unmanagedSheet.Cells["T1:U1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                unmanagedSheet.Cells["T1:U1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["A2:A" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["R2:R" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["W2:W" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
                                unmanagedSheet.Cells[unmanagedSheet.Dimension.Address].AutoFitColumns();
                            }

                            excelPackage.Save();
                            return;
                        }
                        catch (Exception exportEx)
                        {
                            string str = exportEx.Message;
                            return;
                        }
                    }
                }
            }
        }

        public static void GenerateSupplierBuyerReportsSO()
        {
            using (var rdb = new ReportDbEntities())
            {
                //var owners = rdb.BOMShortageProductGroups.Include(x => x.BOMShortageOwners).ToList();
                List<OPSSUPBuyerShortageMetric> savedMetrics = new List<OPSSUPBuyerShortageMetric>();
                //var supplierBuyers = rdb.BOMSupplierManageMatrices.ToList();
                var supplierBuyers = new List<BOMSupplierManageMatrix>();
                using (var connect = new ConnectDbEntities()) { supplierBuyers = connect.BOMSupplierManageMatrices.ToList(); }

                List<Buyer> buyerInfo = new List<Buyer>();
                string theDateHours = DateTime.Now.ToString("yyyyMMdd HH.mm.ss");

                FileInfo fileInfo;
                if (CreateDirectoryStructure(out fileInfo, theDate, theDateHours, @"Shortages_so_BuyerWorkToList", "Shortage Reports", true))  
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                    {
                        var workSheet = excelPackage.Workbook.Worksheets.Add("Buyers Overview");
                        var resultSet = new List<BOMShortagesSupplierBreakdownSO>().ToList();
                        try
                        {
                            rdb.Database.CommandTimeout = 72000;
                            resultSet = rdb.BOMShortagesSupplierBreakdownSOes.ToList();
                        }
                        catch (Exception ex)
                        {
                            //return View();               
                        }
                        try
                        {
                            supplierBuyers.GroupBy(x => x.BuyerName).ToList().ForEach(buyer =>
                            {
                                var buyerSheet = excelPackage.Workbook.Worksheets.Add(buyer.First().BuyerName);
                                var suppliers = buyer.Select(b => b.Supplier).ToList();
                                var supplierShorts = resultSet.Where(r => suppliers.Contains(r.SupplierName)).ToList();

                                List<BOMShortageSupplierBreakdownSOObj> exports2 = new List<BOMShortageSupplierBreakdownSOObj>();
                                List<string> partsNotOnOrder = new List<string>();
                                List<string> partsNotSupported = new List<string>();
                                int notOnOrderCount = 0;
                                int notSupportiveCount = 0;
                                decimal notOnOrderQty = 0;
                                decimal notSupportiveQty = 0;
                                if (supplierShorts.Count() > 0)
                                {
                                    supplierShorts.ForEach(check =>
                                    {
                                        check.ProductGroup = check.ProductGroup == null ? string.Empty : check.ProductGroup;
                                        var export = new BOMShortageSupplierBreakdownSOObj();
                                        export.SalesOrderTitle = check.SalesOrderTitle;
                                        export.DespatchDate = check.DespatchDate.Value.AddDays(42);
                                        export.PODD = check.DespatchDate;
                                        export.ComponentPart = check.ComponentPart;
                                        export.ComponentPartDescription = check.ComponentPartDescription;
                                        export.ComponentMethod = check.ComponentMethod;
                                        export.ProductGroup = check.ProductGroup;
                                        export.Responsibility = check.Responsibility;
                                        export.LeadTime = check.LeadTime.Value;
                                        export.SupplierName = check.SupplierName;
                                        export.Quantity = check.Quantity.GetValueOrDefault();
                                        export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                        export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                        export.Stock = check.Stock.GetValueOrDefault();
                                        export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                        export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                        export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                        export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                        export.POArriving = check.POArriving.GetValueOrDefault();
                                        export.PODelayInDays = check.PODelayInDays;
                                        //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                        //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                        decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                        decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                        decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                        //decimal normalisedShortage = Math.Abs(Shortage) - WOIssued;
                                        //decimal normalisedOnTime = Math.Abs(OnTime) - WOIssued;
                                        decimal normalisedShortage = Shortage + WOIssued;
                                        decimal normalisedOnTime = OnTime + WOIssued;
                                        //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                        //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                        export.Shortage = normalisedShortage;
                                        export.OnTime = normalisedOnTime;
                                        export.WOIssued = WOIssued;
                                        export.UnitCost = check.UnitCost.Value;
                                        export.ReportDate = DateTime.Now.Date.ToShortDateString();
                                        export.BuyerName = buyer.First().BuyerName;
                                        notOnOrderCount += export.Shortage < 0 ? 1 : 0;
                                        notSupportiveCount += export.Shortage >= 0 && export.OnTime < 0 ? 1 : 0;
                                        notOnOrderQty += export.Shortage < 0 ? export.Shortage.Value : 0;
                                        notSupportiveQty += export.Shortage >= 0 && export.OnTime < 0 ? export.OnTime.Value : 0;
                                        if (export.Shortage < 0 && !partsNotOnOrder.Contains(export.ComponentPart))
                                        {
                                            partsNotOnOrder.Add(export.ComponentPart);
                                        }
                                        if (export.Shortage >= 0 && export.OnTime < 0 && !partsNotSupported.Contains(export.ComponentPart))
                                        {
                                            partsNotSupported.Add(export.ComponentPart);
                                        }
                                        if (export.OnTime < 0)
                                            exports2.Add(export);
                                    });

                                    var countz = 2;
                                    foreach (var shortage in exports2)
                                    {
                                        var shortageValue = shortage.Shortage;
                                        decimal dShortValue = Convert.ToDecimal(shortageValue);
                                        var onTimeValue = shortage.OnTime;
                                        decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);

                                        if (dShortValue >= 0)
                                        {
                                            buyerSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                            buyerSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                        }
                                        else
                                        {
                                            buyerSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            buyerSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                        }
                                        if (dOnTimeValue >= 0)
                                        {
                                            buyerSheet.Cells["V" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["V" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                            buyerSheet.Cells["V" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                        }
                                        else
                                        {
                                            buyerSheet.Cells["V" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["V" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            buyerSheet.Cells["V" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                        }
                                        if (shortage.PODD.Value.AddDays(-shortage.LeadTime) < DateTime.Now)
                                        {
                                            buyerSheet.Cells["I" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            buyerSheet.Cells["I" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            buyerSheet.Cells["I" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                        }
                                        countz++;
                                    }

                                    int rowCount = exports2.Count();
                                    string dateTimeFormat = "dd/MM/yyyy";
                                    buyerSheet.Cells["A1:Z1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    buyerSheet.Cells["A1:Z1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                    buyerSheet.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                                    buyerSheet.Cells["A1:I1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["K1:M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                    buyerSheet.Cells["K1:M1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["N1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                    buyerSheet.Cells["N1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["O1:P1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                                    buyerSheet.Cells["O1:P1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    buyerSheet.Cells["Q1:T1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                                    buyerSheet.Cells["Q:T1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    buyerSheet.Cells["U1:V1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                    buyerSheet.Cells["U1:V1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    buyerSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["C2:C" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["S2:S" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["X2:X" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                    buyerSheet.Cells["A1"].LoadFromCollection(exports2, true, OfficeOpenXml.Table.TableStyles.Medium2);
                                    buyerSheet.Cells[buyerSheet.Dimension.Address].AutoFitColumns();
                                };
                                buyerInfo.Add(new Buyer { BuyerName = buyer.First().BuyerName, SupplierCount = suppliers.Count(), ShortageCount = notOnOrderCount + notSupportiveCount, NotOnOrderCount = notOnOrderCount, NotSupportiveCount = notSupportiveCount, NotOnOrderQty = Math.Abs(notOnOrderQty), NotSupportiveQty = Math.Abs(notSupportiveQty), UniquePartsNotOnOrder = partsNotOnOrder.Count(), UniquePartsNotSupportingPODD = partsNotSupported.Count() });
                                savedMetrics.Add(new OPSSUPBuyerShortageMetric { BuyerName = buyer.First().BuyerName, SupplierCount = suppliers.Count(), ShortageCount = notOnOrderCount + notSupportiveCount, NotOnOrder = notOnOrderCount, NotSupportingPODD = notSupportiveCount, RecordingDate = DateTime.Now.Date, NotOnOrderQty = Math.Abs(notOnOrderQty), NotSupportingPODDQty = Math.Abs(notSupportiveQty), UniquePartsNotOnOrder = partsNotOnOrder.Count(), UniquePartsNotSupportingPODD = partsNotSupported.Count() });
                            });

                            // Group unmanaged suppliers together...
                            var unmanagedSheet = excelPackage.Workbook.Worksheets.Add("Unassigned Suppliers");
                            var supplierShorts2 = resultSet.Where(r => !supplierBuyers.Select(sup => sup.Supplier).ToList().Contains(r.SupplierName)).ToList();
                            buyerInfo.Add(new Buyer { BuyerName = "Unassigned Suppliers", SupplierCount = supplierShorts2.Select(um => um.SupplierName).ToList().Distinct().Count(), ShortageCount = supplierShorts2.Count() });
                            workSheet.Cells["A1"].LoadFromCollection(buyerInfo, true, OfficeOpenXml.Table.TableStyles.Medium2);
                            workSheet.Cells["A1:G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells["A1:G1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                            workSheet.Cells["A1"].Value = "Buyer Name";
                            workSheet.Cells["B1"].Value = "Supplier";
                            workSheet.Cells["C1"].Value = "Shortages";
                            workSheet.Cells["D1"].Value = "Lines Not On Order";
                            workSheet.Cells["E1"].Value = "Quantity Not On Order";
                            workSheet.Cells["F1"].Value = "Unique Part Count Not On Order";
                            workSheet.Cells["G1"].Value = "Lines Not Supporting PODD";
                            workSheet.Cells["H1"].Value = "Quantity Not Supporting PODD";
                            workSheet.Cells["I1"].Value = "Unique Part Count Not Supporting PODD";
                            workSheet.Column(1).Width = 20;
                            workSheet.Column(2).Width = 15;
                            workSheet.Column(3).Width = 15;
                            workSheet.Column(4).Width = 20;
                            workSheet.Column(5).Width = 25;
                            workSheet.Column(6).Width = 32.5;
                            workSheet.Column(7).Width = 32.5;
                            workSheet.Column(8).Width = 32.5;
                            workSheet.Column(9).Width = 32.5;

                            workSheet.Column(4).Hidden = true;
                            workSheet.Column(5).Hidden = true;
                            workSheet.Column(6).Hidden = true;
                            workSheet.Column(7).Hidden = true;

                            List<BOMShortageSupplierBreakdownObj> exports = new List<BOMShortageSupplierBreakdownObj>();
                            decimal? priority = new decimal(0.0);
                            if (supplierShorts2.Count() > 0)
                            {
                                supplierShorts2.ForEach(check =>
                                {
                                    check.ProductGroup = check.ProductGroup == null ? string.Empty : check.ProductGroup;
                                    var export = new BOMShortageSupplierBreakdownObj();
                                    export.DespatchDate = check.DespatchDate.Value.AddDays(42);
                                    export.PODD = check.DespatchDate;
                                    export.ComponentPart = check.ComponentPart;
                                    export.ComponentPartDescription = check.ComponentPartDescription;
                                    export.ComponentMethod = check.ComponentMethod;
                                    export.ProductGroup = check.ProductGroup;
                                    export.Responsibility = check.Responsibility;
                                    export.LeadTime = check.LeadTime.Value;
                                    export.SupplierName = check.SupplierName;
                                    export.Quantity = check.Quantity.GetValueOrDefault();
                                    export.PriorDemand = check.PriorDemand.GetValueOrDefault();
                                    export.TotalDemand = check.TotalDemand.GetValueOrDefault();
                                    export.Stock = check.Stock.GetValueOrDefault();
                                    export.WoQuantity = check.WoQuantity.GetValueOrDefault();
                                    export.WoOnTime = check.WoOnTime.GetValueOrDefault();
                                    export.PoQuantity = check.PoQuantity.GetValueOrDefault();
                                    export.PoOnTime = check.PoOnTime.GetValueOrDefault();
                                    export.POArriving = check.POArriving.GetValueOrDefault();
                                    export.PODelayInDays = check.PODelayInDays;
                                    //export.Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  //(check.Quantity.GetValueOrDefault() + priority.GetValueOrDefault());             
                                    //export.OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());  // check.Quantity.GetValueOrDefault();
                                    decimal Shortage = (check.Stock.GetValueOrDefault() + check.WoQuantity.GetValueOrDefault() + check.PoQuantity.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal OnTime = (check.Stock.GetValueOrDefault() + check.WoOnTime.GetValueOrDefault() + check.PoOnTime.GetValueOrDefault()) - (check.Quantity.GetValueOrDefault() + check.PriorDemand.GetValueOrDefault());
                                    decimal WOIssued = check.WOIssued.GetValueOrDefault();
                                    //decimal normalisedShortage = Math.Abs(Shortage) - WOIssued;
                                    //decimal normalisedOnTime = Math.Abs(OnTime) - WOIssued;
                                    decimal normalisedShortage = Shortage + WOIssued;
                                    decimal normalisedOnTime = OnTime + WOIssued;
                                    //export.Shortage = normalisedShortage < 0 ? 0 : normalisedShortage;
                                    //export.OnTime = normalisedOnTime < 0 ? 0 : normalisedOnTime;
                                    export.Shortage = normalisedShortage;
                                    export.OnTime = normalisedOnTime;
                                    export.WOIssued = WOIssued;
                                    export.UnitCost = check.UnitCost.Value;
                                    export.ReportDate = DateTime.Now.Date.ToShortDateString();
                                    exports.Add(export);
                                });

                                var countz = 2;
                                foreach (var shortage in exports)
                                {
                                    var shortageValue = shortage.Shortage;
                                    decimal dShortValue = Convert.ToDecimal(shortageValue);
                                    var onTimeValue = shortage.OnTime;
                                    decimal dOnTimeValue = Convert.ToDecimal(onTimeValue);

                                    if (dShortValue >= 0)
                                    {
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        unmanagedSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["U" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        unmanagedSheet.Cells["U" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    if (dOnTimeValue >= 0)
                                    {
                                        unmanagedSheet.Cells["V" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["V" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        unmanagedSheet.Cells["V" + countz].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    }
                                    else
                                    {
                                        unmanagedSheet.Cells["V" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["V" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        unmanagedSheet.Cells["V" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    if (shortage.PODD.Value.AddDays(-shortage.LeadTime) < DateTime.Now)
                                    {
                                        unmanagedSheet.Cells["I" + countz].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        unmanagedSheet.Cells["I" + countz].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                        unmanagedSheet.Cells["I" + countz].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                    }
                                    countz++;
                                }

                                int rowCount = exports.Count();
                                string dateTimeFormat = "dd/MM/yyyy";
                                unmanagedSheet.Cells["A1:Z1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                unmanagedSheet.Cells["A1:Z1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                unmanagedSheet.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
                                unmanagedSheet.Cells["A1:I1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["K1:M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                                unmanagedSheet.Cells["K1:M1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["N1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                unmanagedSheet.Cells["N1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["O1:P1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Goldenrod);
                                unmanagedSheet.Cells["O1:P1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                unmanagedSheet.Cells["Q1:T1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGoldenrod);
                                unmanagedSheet.Cells["Q:T1"].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                unmanagedSheet.Cells["U1:V1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                unmanagedSheet.Cells["U1:V1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                                unmanagedSheet.Cells["B2:B" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["C2:C" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["S2:S" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["X2:X" + rowCount].Style.Numberformat.Format = dateTimeFormat;
                                unmanagedSheet.Cells["A1"].LoadFromCollection(exports, true, OfficeOpenXml.Table.TableStyles.Medium2);
                                unmanagedSheet.Cells[unmanagedSheet.Dimension.Address].AutoFitColumns();
                            }

                            excelPackage.Save();
                            CopyBuyerMetricsToDB(savedMetrics);
                            return;
                        }
                        catch (Exception exportEx)
                        {
                            string str = exportEx.Message;
                            string trace = exportEx.StackTrace;
                            return;
                        }
                    }
                }
            }
        }

        //private static void SaveBuyerMetricsToDatabase(List<OPSSUPBuyerShortageMetric> metrics)
        //{
        //    using (var rdb = new ReportDbEntities())
        //    {
        //        rdb.OPSSUPBuyerShortageMetrics.AddRange(metrics);
        //        rdb.SaveChanges();
        //    }
        //}

        public static void CopyBuyerMetricsToDB(List<OPSSUPBuyerShortageMetric> dataSet)
        {
            ReportDbEntities connect = null;
            try
            {
                connect = new ReportDbEntities();
                connect.Configuration.AutoDetectChangesEnabled = false;

                int count = 0;
                foreach (var line in dataSet)
                {
                    ++count;
                    connect = AddToContextBuyerMetrics(connect, line, count, 500, true);
                }
                connect.SaveChanges();
            }
            finally
            {
                if (connect != null)
                    connect.Dispose();
            }
        }

        private static ReportDbEntities AddToContextBuyerMetrics(ReportDbEntities context, OPSSUPBuyerShortageMetric entity, int count, int commitCount, bool recreateContext)
        {
            context.Set<OPSSUPBuyerShortageMetric>().Add(entity);

            if (count % commitCount == 0)
            {
                context.SaveChanges();
                if (recreateContext)
                {
                    context.Dispose();
                    context = new ReportDbEntities();
                    context.Configuration.AutoDetectChangesEnabled = false;
                }
            }
            return context;
        }

        private static void GenerateSalesOrdersWithoutBOMsReport(List<string> salesOrders)
        {
            FileInfo fileInfo;
            string theDate = targetDate.ToString("yyyyMMdd");
            string theDateHours = targetDate.ToString("yyyyMMdd HH.mm.ss");

            if (CreateDirectoryStructure(out fileInfo, theDate, theDateHours, "Sales-Orders-Missing-BOMs-Report"))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    var workSheet = excelPackage.Workbook.Worksheets.Add("SOs Missing BOMs");
                    List<SOMissingBOM> missingBOMs = new List<BomExcelGenerator.SOMissingBOM>();
                    salesOrders.ForEach(x =>
                    {
                        missingBOMs.Add(new SOMissingBOM { SalesOrder = x });
                    });

                    int rowCount = missingBOMs.Count();
                    workSheet.Cells["A1:A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells["A1:A1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue);
                    workSheet.Cells["A1"].LoadFromCollection(missingBOMs, true, OfficeOpenXml.Table.TableStyles.Medium2);
                    workSheet.Column(1).Width = 50;
                    excelPackage.Save();
                }
            }
        }

        private static BOMLineShortage GetShortageForDB(BOMExport export)
        {
            return new BOMLineShortage
            {
                SalesOrderTitle = export.SalesOrderTitle,
                DespatchDate = export.PODD,
                CustReqDate = export.CustReqDate,
                MainPart = export.MainPart,
                MainPartDescription = export.MainPartDescription,
                ComponentPart = export.ComponentPart,
                ComponentPartDescription = export.ComponentPartDesc,
                ComponentMethod = export.ComponentMethod,
                Responsibility = export.Responsibility,
                ProductGroup = export.ProductGroup,
                ResourceType = export.ResourceType,
                ResourceCode = export.ResourceCode,
                ResourceGroupName = export.ResourceGroupName,
                UnitOfMeasure = export.UnitOfMeasure,
                Quantity = export.Quantity,
                TotalBOMQuantity = export.TotalBOMQuantity,
                PriorDemand = export.PriorDemand,
                TotalDemand = export.TotalDemand,
                Stock = export.Stock,
                WoQuantity = export.WoQuantity,
                WoOnTime = export.WoOnTime,
                WoArriving = export.WoArriving,
                WoDelayInDays = export.WoDelayInDays,
                PoQuantity = export.PoQuantity,
                PoOnTime = export.PoQuantity,
                PoArriving = export.PoArriving,
                PoDelayInDays = export.PoDelayInDays,
                Shortage = export.Shortage,
                OnTime = export.OnTime
            };
        }

        private static bool CreateDirectoryStructure(string milestone, out FileInfo fileInfo, string date)
        {
            fileInfo = new FileInfo(string.Format(@"\\thas-report01\ShortageReports\{0}\{1}_{0}.xlsx", date, milestone));
            try
            {
                var fullpath = string.Format(@"\\thas-report01\ShortageReports\{0}\{1}_{0}.xlsx", date, milestone);
                if (!File.Exists(fullpath))
                {
                    fileInfo = new FileInfo(fullpath);
                    fileInfo.Directory.Create();
                    return true;
                }
                else
                    return false; // get out of here.              
            }
            catch (Exception ex)
            {
                Console.WriteLine("Issue : " + ex.Message);
                return false;
            }
        }

        private static bool CreateDirectoryStructure(out FileInfo fileInfo, string date, string dateHours, string filename, string folderPath, bool costed)
        {
            string path = @"\\tas\reports$\{0}\{1}\";
            if (costed)
            {
                path = @"\\tas\reports$\{0}\With Costing Info\{1}\";
            }
            else
            {
                path = @"\\tas\reports$\{0}\Without Costing Info\{1}\";
            }

            fileInfo = new FileInfo(string.Format(path + filename + "_{2}.xlsx", folderPath, date, dateHours));
            try
            {
                var fullpath = string.Format(path + filename + "_{2}.xlsx", folderPath, date, dateHours);
                if (!File.Exists(fullpath))
                {
                    fileInfo = new FileInfo(fullpath);
                    fileInfo.Directory.Create();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Issue : " + ex.Message);
                return false;
            }
        }

        private static bool CreateDirectoryStructure(out FileInfo fileInfo, string date, string dateHours, string filename)
        {
            string path = @"G:\TAS Perform\TAS Daily Metrics\{0}\";
            fileInfo = new FileInfo(string.Format(path + filename + "_{1}.xlsx", date, dateHours));
            try
            {
                //var fullpath = string.Format(@"S:\OpsSupport\DailyMetrics\{0}\PartAvailabilityMetrics_{1}.xlsx", date, dateHours);
                var fullpath = string.Format(path + filename + "_{1}.xlsx", date, dateHours);
                if (!File.Exists(fullpath))
                {
                    fileInfo = new FileInfo(fullpath);
                    fileInfo.Directory.Create();
                    return true;
                }
                else
                    return false; // get out of here.              
            }
            catch (Exception ex)
            {
                Console.WriteLine("Issue : " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static List<string> GetScheduleOfSalesOrders()
        {
            List<string> so_range = new List<string>();
            DateTime today = DateTime.Now;
            DateTime start = today; // It used to be 2 weeks to 12 weeks out. - for now it is the full 12 weeks.
            //DateTime start = today.AddDays(14);
            DateTime end = today.AddDays(84); // .AddDays(84)
            Console.WriteLine("----- BEGIN Calculating Reporting Period -----");
            Console.WriteLine("Reporting Period Start : " + start.ToShortDateString());
            Console.WriteLine("Reporting Period End : " + end.ToShortDateString());
            using (var lol = new thas01Entities())
            {
                try
                {
                    so_range = lol.THAS_CONNECT_GetShortageReportSOSchedule(start, end).ToList();
                    Console.WriteLine("Found " + so_range.Count.ToString() + " Sales Orders within the Reporting Period.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("***** ERROR Calculating Reporting Period.  Details : " + ex.Message);
                }
            }
            Console.WriteLine("----- FINISHED Calculating Reporting Period -----");
            return so_range;
        }


        private static void PrintWelcomeMessage()
        {
            Console.WriteLine("");
            Console.WriteLine(" --------------------------------------------- ");
            Console.WriteLine("");
            Console.Title = " Full BOM Generator (x64) [" + typeof(BomExcelGenerator.Program).Assembly.GetName().Version + "]";
            Console.WriteLine(" Welcome to the TAS Connect BOM Generator");
            Console.WriteLine(" Version : " + typeof(BomExcelGenerator.Program).Assembly.GetName().Version);
            Console.WriteLine("");
            Console.WriteLine(" --------------------------------------------- ");
            Console.WriteLine("");
            Console.WriteLine(" Here are your options...");
            Console.WriteLine("");
            Console.WriteLine(" --------------------------------------------- ");
            Console.WriteLine("");
            Console.WriteLine(" 1. Run Shortage Builder, Re-Compute & Milestone Shortage Report Exporter.");
            Console.WriteLine("");
            Console.WriteLine(" 2. Run BOM Computation Only.");
            Console.WriteLine("");
            Console.WriteLine(" 3. Process Shortages & Run Milestone Shortage Report Exporter.");
            Console.WriteLine("");
            Console.WriteLine(" 4. Run Milestone Shortage Report Exporter.");
            Console.WriteLine("");
            Console.WriteLine(" 5. Run TEST INDIVID Milestone Processor.");
            Console.WriteLine("");
            Console.WriteLine(" 6. Process Giants.");
            Console.WriteLine("");
            Console.WriteLine(" 7. Perform Supplier Breakdown Report & Export.");
            Console.WriteLine("");
            Console.WriteLine(" 0. Run on Auto Timer (8pm Run)");
            Console.WriteLine("");
            Console.WriteLine(" --------------------------------------------- ");
            Console.WriteLine("");

            AppLogger.ReportInfo("Starting Shortage Report Generator.");
        }


    }
}
