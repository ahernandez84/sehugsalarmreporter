using System;
using System.Data;
using System.Drawing;
using System.Configuration;
using System.IO;
using System.Linq;

using NLog;
using SeHugsAlarmReporter.Models;
using SeHugsAlarmReporter.Services;

namespace SeHugsAlarmReporter
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("<< SE Hugs Alarm Reporter v1.0.0 >>");

            logger.Info("SE Hugs Alarm Reporter is initializing.");

            var appSettings = ReadApplicationConfig();

            if (!CreateReportDirectory(appSettings.FilePath))
            {
                logger.Warn("Failed to create the report directory.");
                return;
            }

            //Get connection string
            var connectionString = @"Server=(local)\xmark; database=xmark; trusted_connection=true;";

            logger.Info("Attempting to get Hugs Counts.");

            //Generate counts
            var deviceResults = SQLService.GetSystemCounts(connectionString);
            if (deviceResults == null)
            {
                Console.WriteLine("Failed to find device counts.  See log for more information.");
                logger.Warn("Failed to find device counts.");
                return;
            }

            var alarmResults = SQLService.GetAlarmCounts(connectionString);
            if (alarmResults == null)
            {
                Console.WriteLine("Failed to find alarm counts.  See log for more information.");
                logger.Warn("Failed to find alarm counts.");
                return;
            }

            alarmResults = ModifyAlarmResults(alarmResults);

            var filenamePath = $@"{appSettings.FilePath}\{DateTime.Today.AddMonths(-1).ToString("MMMM yyyy")} Hugs Report";
            var reportName = $"{DateTime.Today.AddMonths(-1).ToString("MMMM yyyy")} Hugs Report";

            logger.Info($"Attempting to generate Excel report:  {filenamePath}");

            var result = ExcelService.GenerateReport(filenamePath, reportName, new[] { deviceResults, alarmResults }, appSettings.ReportColor);

            if (result && appSettings.SendEmail)
                EmailSingletonService.Instance.SendEmail(filenamePath + ".xlsx");
        }

        #region Private Methods
        private static DataTable ModifyAlarmResults(DataTable dt)
        {
            var alarmTypes = XMLService.LoadFile();

            dt.Columns.Add("Critical", typeof(string));
            dt.Columns.Add("Note", typeof(string));

            foreach (DataRow dr in dt.Rows)
            {
                var alarmType = alarmTypes.Where(w => dr[0].ToString().ToUpper().Contains(w.Name.ToUpper())).SingleOrDefault();

                dr[0] = alarmType == null ? dr[0].ToString() : alarmType.Name;
                dr["Critical"] = alarmType == null ? "" : alarmType.Critical;
                dr["Note"] = alarmType == null ? "" : alarmType.Note;
            }

            return dt;
        }

        private static AppSettings ReadApplicationConfig()
        {
            try
            {
                var settings = new AppSettings();
                settings.FilePath = ConfigurationManager.AppSettings["Filepath"];
                settings.SendEmail = Convert.ToBoolean(ConfigurationManager.AppSettings["SendEmail"]);

                var color = ConfigurationManager.AppSettings["ReportColor"];
                settings.ReportColor = Color.FromArgb(Convert.ToInt32(color.Split('.')[0]), Convert.ToInt32(color.Split('.')[1]), Convert.ToInt32(color.Split('.')[2]));

                return settings;
            }
            catch (Exception ex) { logger.Error(ex, "Program <ReadApplicationConfig> method."); return null; }
        }

        private static bool CreateReportDirectory(string filePath)
        {
            try
            {
                if (!Directory.Exists(filePath))
                {
                    var dirInfo = Directory.CreateDirectory(filePath);
                    dirInfo.Refresh();
                }

                return true;
            }
            catch (Exception ex) { logger.Error(ex, "Program <CreateReportDirectory> method."); return false; }
        }
        #endregion

    }
}
