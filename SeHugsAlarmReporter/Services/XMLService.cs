using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

using SeHugsAlarmReporter.Models;

namespace SeHugsAlarmReporter.Services
{
    internal class XMLService
    {
        //private static Logger logger = LogManager.GetCurrentClassLogger();
        public static List<AlarmType> LoadFile()
        {
            try
            {
                //logger.Info($"App Directory:  {AppDomain.CurrentDomain.BaseDirectory}");

                XDocument doc = XDocument.Load($@"{AppDomain.CurrentDomain.BaseDirectory}\AlarmType.xml");

                var types = from p in doc.Descendants("Type")
                             select new AlarmType()
                             {
                                 Name =p.Element("Name").Value,
                                 Critical = p.Element("Critical").Value,
                                 Note = p.Element("Note").Value
                             };

                //logger.Info($"XML Service found {points.Count()} point(s) to monitor.");

                return types.ToList<AlarmType>();
            }
            catch (Exception ex) { //logger.Error(ex, "XMLService <LoadFile> method."); 
                return null;
            }
        }
    }
}
