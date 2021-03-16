using System;
using System.Drawing;
namespace SeHugsAlarmReporter.Models
{
    class AppSettings
    {
        public bool SendEmail { get; set; }
        public string FilePath { get; set; }
        public Color ReportColor { get; set; }
    }
}
