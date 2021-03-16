using System;
using System.Data;
using System.Data.SqlClient;

using NLog;

namespace SeHugsAlarmReporter.Services
{
    class SQLService
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static DataTable GetSystemCounts(string connectionString)
        {
            #region Query
            var query = @"select 'Local Area Receivers (LARS)' as 'Device Type', COUNT(*) as 'Count', '', ''
                          from tblReceiver
                          union
                          select 'Portal Exciters' as 'Device Type', COUNT(*) as 'Count', '', ''
                          from tblExciter
                          union
                          select 'Workstations' as 'Device Type', COUNT(*) as 'Count', '', ''
                          from tblPC
                          union
                          select 'Communication Panels' as 'Device Type', 20 as 'Count', '', ''
                          from tbllon
                          --union
                          --select 'System Alarm Testing' as 'Device Type',90 as 'Count', '', ''
                          --from tbllon";
            #endregion

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        var results = new DataTable("SystemCounts");

                        adapter.Fill(results);

                        return results;
                    }
                }
            }
            catch (Exception ex) { logger.Error(ex, "SQLService <GetSystemCounts> method."); Console.WriteLine($"SQLService: {ex.Message}"); return null; }
        }

        public static DataTable GetAlarmCounts(string connectionString)
        {
            #region Query
            var query = @";with alarms as (
	                        select distinct 
		                        a.DateTime
		                        ,b.ID as 'TagId'
		                        ,g.Description as 'EventType' 
		                        ,c.Description as 'Exciter'
		                        ,b.description as 'TagName'
		                        ,a.EventDetail as 'TagDetails'
		                        --,a.EventTypeKey
		                        --,z.[description] as 'TagZone'
	                        from [xmark].dbo.tblEventLog a with(nolock)
	                        left outer join [xmark].dbo.tblEventObject g with(nolock) on a.EventKey=g.AutoID  
	                        left outer join [xmark].dbo.tblEventObject h with(nolock) on a.EventTypeKey=h.AutoID 
	                        left outer join [xmark].dbo.tblEventObject b with(nolock) on a.ItemKey=b.AutoID 
	                        left outer join [xmark].dbo.tblEventObject d with(nolock) on a.UserKey=d.AutoID 
	                        left outer join [xmark].dbo.tblEventObject c with(nolock) on a.LocationKey=c.AutoID 
	                        left join [xmark].dbo.tblbabytag t with(nolock) on (t.id = b.id)
	                        --left join [xmark].dbo.tblzone z with(nolock) on (z.id = t.[zone])
	                        where 
		                        a.datetime >= @start and a.DateTime <= @end
		                        and a.EventTypeKey in (4097,4101)
		                        and (b.Type=21) 
		                        and (g.description like 'Admit Acknowledgement Alarm%' or g.description like 'Improperly Applied Tag%'  or g.description like 'Band Detached Alarm%' 
			                        or g.description like 'Tag Exit Alarm%' or g.description like 'Tag loose - tighten immediately%' or g.description like 'Tag supervision timeout%' 
			                        or g.description like 'Tag Tamper%' or g.description like 'Auto Discharge%' or g.description like 'check tag%') 
		                        and b.Description not like '%schneider%'
	                        --order by g.Description, a.DateTime asc
                        )
                        select a.EventType as 'Alarm Type', COUNT(*) as 'Total Alarms'
                        from alarms a
                        group by a.EventType
                        order by a.EventType";
            #endregion

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        var start = new DateTime(DateTime.Now.Year, DateTime.Now.AddMonths(-1).Month, 1);
                        var end = start.AddDays(DateTime.DaysInMonth(start.Year, start.Month) - 1);

                        command.Parameters.AddWithValue("@start", start.ToString("MM/dd/yyyy"));
                        command.Parameters.AddWithValue("@end", end.ToString("MM/dd/yyyy"));

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            var results = new DataTable("AlarmCounts");

                            adapter.Fill(results);

                            return results;
                        }
                    }
                }
            }
            catch (Exception ex) { logger.Error(ex, "SQLService <GetAlarmCounts> method."); Console.WriteLine($"SQLService: {ex.Message}"); return null; }
        }


    }
}
