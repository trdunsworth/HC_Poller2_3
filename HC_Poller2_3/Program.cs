using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;

namespace HC_Poller2_3
{
  /// <summary>
  /// Hotcalls, currently, is a four-piece application which monitors the JCSOARCH database for events which 
  /// reach a collections of triggers (Location, Call Type & Subtype, or number of arrived units).
  /// Two parts are the poller which queries the archive database and consolodates that data into
  /// tables created for this program for evaluation. The poller also takes care of removing rows from that database
  /// based on age of the call or other factors. The logic program evaluates the daata and then if a trigger is matched,
  /// will then send out an email to subscribed users using the county Exchange Server. The third component is an
  /// administrative component which allows administrators to add and modify users, add and modify locations, and 
  /// monitor what is being generated as well as making minor adjustments, on the fly, to those parameters.
  /// The fourth component is a health monitor which runs to determine when if the processes have stopped and notifies the administrators as needed.
  /// The poller is started first and the logic portion runs approximately 60 seconds later. 
  /// Both programs are run on Windows Server 2008 by the Task Scheduler on a 2 minute rotation roughly 60 seconds apart
  /// Both are installed and scheduled via the local admin account.
  /// The poller and logic program executables live in the following path:
  /// C:\Users\jcsoadmin\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Poller -or- Logic
  /// Those will now change as of Version 2 to C:\Program Files\ with the use of the Inno Script setup program
  /// All errors are collected and written to a log file in the C:\Temp folder.
  /// Both programs also email the administrators to enable rapid response 
  /// Created in Visual Studio 2013 and using the C# 5.0 Language specification
  /// Copyright 2014, Johnson County Sheriff's Office. All rights reserved.
  /// </summary>
  // Program: Hot Calls Poller Version 2 Write 3
  // Author: Tony Dunsworth
  // Date Created: 20140813 (Version 2.0.3.0001)
  // Happy Birthday to Ellen Dunsworth edition.
  // Date of current revision: 20140813
  // Date of TODO List 2014-01-09
  // 1. Reconfigure db tables where appropriate to use EF versions more efficiently
  // 2. Continue to monitor output for possible table change for unit arrival counts.
  // 3. Condense multiple queries into one method for better handling.
  class Program
  {
    // Creates the database connection string for use in the program.
    public static string qry;
    public static OracleConnection cxn = new OracleConnection(ConfigurationManager.ConnectionStrings["hotcall"].ConnectionString);

    // The Main function contains the program flow, invoking a legacy timer if needed, which would run every 120 seconds raising an Elapsed event on fire
    // It also provides for a graceful shutdown for maintenance with the enter key (when using the legacy timer)
    // Finally, there was a garbage collector for memory and resource management.
    static void Main(string[] args)
    {
      // The function starts with a time stamp on the console to give visual verification everything is running as scheduled.
      // Any freeze in the process can provide a visual clue that the program is failing.
      // Changed the string from a custom format to a .NET supplied format to reduce potential error points.
      Console.WriteLine("Poller entering at " + DateTime.Now.ToLongTimeString());
      CollectCurrentCalls();
      CollectCurrentComments();
      CollectCurrentCount();
      WriteCurrentCalls();
      WriteCurrentComments();
      WriteCurrentCount();
      PurgeCurrentCalls();
      PurgeCurrentComments();
      PurgeCurrentCount();
      CleanHotCalls();
      DeleteHotCalls();
      Console.WriteLine("Poller exiting at " + DateTime.Now.ToLongTimeString());
    }

    private static void CollectCurrentCalls()
    {
      // Each function will begin with a time stamp on the console for debugging when necessary
      // In live code this is not necessary unless there is a failure.
      // Console.WriteLine("Collect Current Calls entering at " + DateTime.Now.ToLongTimeString());
      // The query string will collect base call data from the agency_event and common_event tables.
      // The entry time must be in the last 8 minutes and the call must be open to be selected.
      qry = "INSERT INTO hc_curent_temp(eid, ag_id, tycod, sub_tycod, ad_ts, udts, xdts, num_1, estnum, edirpre, efeanme, efeatyp, xstreet1, xstreet2, esz) SELECT DISTINCT a.eid, a.ag_id, a.tycod, a.sub_tycod, a.ad_ts, a.udts, a.num_1, c.estnum, c.edirpre, c.efeanme, c.efeatyp, c.xstreet1, c.xstreet2, a.esz FROM agency_event a JOIN common_event c ON a.eid=c.eid WHERE a.ad_ts > TO_CHAR(systimestamp-8/1440, 'YYYYMMDDHH24MISS')";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      // Added an OracleTransaction for commit and to prevent commits on an error.
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      // Since all of the queries are INSERT, UPDATE, MERGE, or DELETE, we only need ExecuteNonQuery()
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        // Close the connection for security purposes
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The ending Console.WriteLine is for debugging purposes to show when the subroutine exits. This is left out of the production code base
        // Console.WriteLine("Step 1 exiting at " + DateTime.Now.ToString("G");
      }
    }

    private static void CollectCurrentComments()
    {
      // The starting Console.WriteLine statement is used in debugging and testing to show when the program steps into this method.
      // Console.WriteLine("Step 2 entering at " + DateTime.Now.ToString("G");
      // The query uses the agency_event table for calls in the last few minutes then queries the evcom table for the respective comments
      // The "@" at the beginning of the quotation is to ensure the regular expressions will not cause an issue for C#
      // The substring construction is using a custom PL/SQL function to grab the first 4000 characters of the comments.
      qry = @"INSERT INTO hc_comment_temp(eid, num_1, ad_ts, comments) SELECT a.eid, a.num_1, a.ad_ts, substr(TO_NTT(CAST(COLLECT(e.comm ORDER BY e.cdts, e.lin_grp, e.lin_ord) AS varchar2_ntt)),1,4000) FROM agency_event JOIN evcom e ON a.eid=e.eid WHERE a.ad_ts > TO_CHAR(systimestamp-8/1440, 'YYYYMMDDHH24MISS') AND e.comm_key=0 AND NOT(REGEXP_LIKE(e.comm, '^[A-Z,/]{3,}(\s[A-Z]{2,})?(\s[A-Z]{2,)?:') OR REGEXP_LIKE(e.comm, '[A-Z,^0-9]{3,}/') OR REGEXP_LIKE(e.comm, 'END OF (K)?DOR RESPONSE') OR REGEXP_LIKE(e.comm, '-{3,}') OR REGEXP_LIKE(e.comm, '10-[0-9]{2} \*{2,}') OR REGEXP_LIKE(e.comm, 'LICENSE:') OR REGEXP_LIKE(e.comm, '\*{3,}') OR REGEXP_LIKE(e.comm, 'Field Event') OR REGEXP_LIKE(e.comm, '\*{2} Event held for [0,9]{2} minutes')) GROUP BY a.eid, a.num_1, a.ad_ts";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        // Close the connection for security purposes.
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The ending Console.WriteLine statement is for debugging purposes to show when the subroutine exits. This is left out of the production code base
        // Console.WriteLine("Step 2 exiting at " + DateTime.Now.ToString("G");
      }
    }

    // This method is designed to query the agency_event and un_hi tables to aggregate the number of units arrived at an event in the first
    // 25 minutes after call creation.
    private static void CollectCurrentCount()
    {
      // The starting Console.WriteLine statement is for debugging purposes to show when the subroutine enters. This is commented out of the production 
      // code base
      // Console.WriteLine("Step 3 entering at " + DateTime.Now.ToString("G");
      qry = "INSERT INTO hc_unitcount_temp(eid, num_1, ag_id, ad_ts, unit_count) SELECT a.eid, a.num_1, a.ag_id, a.ad_ts, COUNT(DISTINCT u.unid) FROM agency_event a JOIN un_hi u ON (a.eid=u.eid AND a.num_1=u.num_1 AND a.ag_id=u.ag_id) WHERE a.ad_ts>TO_CHAR((systimestamp-25/1440), 'YYYYMMDDHH24MISS') AND u.unit_status='AR' GROUP BY a.eid, a.num_1, a.ag_id, a.ad_ts";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The ending Console.WriteLine statement is for debugging purposes to show when the subroutine exits. This is left out of the production code base
        // Console.WriteLine("Step 3 exiting at " + DateTime.Now.ToString("G");
      }
    }

    // Now that the poller has accumulated the data in the _temp tables, it next needs to add to or update the evaluation tables with this data. As the 
    // subroutine runs, if it encounters a combination of the eid, the num_1, and the ad_ts, it will merge the remaining rows accordingly. However, if it
    // should not find that combination, the subroutine will, through the SQL statement, create a new line in the table for that information.
    private static void WriteCurrentCalls()
    {
      // The starting Console.WriteLine statement is for debugging purposes to show when the subroutine enters. This is commented out of the production 
      // code base
      // Console.WriteLine("Step 4 entering at " + DateTime.Now.ToString("G");
      qry = "MERGE INTO jc_hc_curent j USING (SELECT DISTINCT eid, ag_id, tycod, sub_tycod, ad_ts, udts, xdts, num_1, estnum, edirpre, efeanme, efeatyp, xstreet1, xstreet2, esz FROM hc_curent_temp) h ON (j.eid=h.eid AND j.num_1=h.num_1 AND j.ad_ts=h.ad_ts) WHEN MATCHED THEN UPDATE SET j.udts-h.udts, j.xdts=h.xdts, j.ag_id=h.ag_id, j.tycod=h.tycod, j.sub_tycod-h.sub_tycod, j.estnum=h.estnum, j.edirpre=h.edirpre, j.efeanme=h.efeanme, j.efeatyp=h.efeatyp, j.xstreet1=h.xstreet1, j.xstreet2=h.xstreet2, j.esz=h.esz WHEN NOT MATCHED THEN INSERT(eid, ag_id, tycod, sub_tycod, ad_ts, udts, xdts, num_1, estnum, edirpre, efeanme, efeatyp, xstreet1, xstreet2, esz) VALUES(h.eid, h.ag_id, h.tycod, h.sub_tycod, h.ad_ts, h.udts, h.xdts, h.num_1, h.estnum, h.edirpre, h.efeanme, h.efeatyp, h.xstreet1, h.xstreet2, h.esz)";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The ending Console.WriteLine statement is for debugging purposes to show when the subroutine exits. This is left out of the production code base
        // Console.WriteLine("Step 4 exiting at " + DateTime.Now.ToString("G");
      }
    }

    // This subroutine attempts to merge or insert the comments from the temp table to the evaluation table.
    // If there is not a match, the comments are discarded and the program moves along and cleans up later.
    private static void WriteCurrentComments()
    {
      // The starting Console.WriteLine statement is for debugging purposes to show when the subroutine enters. This is commented out of the production 
      // code base
      // Console.WriteLine("Step 5 entering at " + DateTime.Now.ToString("G");
      qry = "MERGE INTO jc_hc_curent j USING (SELECT DISTINCT eid, num_1, ad_ts, comments FROM hc_comment_temp) c ON (j.eid=c.eid AND j.num_1=c.num_1 AND j.ad_ts=c.ad_ts) WHEN MATCHED THEN UPDATE SET j.comments=c.comments";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      txn.Dispose();
      cmd.Dispose();
      cxn.Close();
      cxn.Dispose();
      // The ending Console.WriteLine statement is for debugging purposes to show when the subroutine exits. This is left out of the production code base
      // Console.WriteLine("Step 5 exiting at " + DateTime.Now.ToString("G");
    }

    // This subroutine take the current units arrived count from the temp table and merges it into the evaluation table. If there is no match here,
    // the data is also discarded here.
    private static void WriteCurrentCount()
    {
      // The starting Console.WriteLine statement is for debugging purposes to show when the subroutine enters. This is commented out of the production 
      // code base
      // Console.WriteLine("Step 6 entering at " + DateTime.Now.ToString("G"));
      qry = "MERGE INTO jc_hc_curent j USING (SELECT DISTINCT eid, ad_ts, num_1, ag_id, unit_count FROM hc_unitcount_temp) u ON (j.eid=u.eid AND j.ad_ts=u.ad_ts AND j.num_1=u.num_1 AND j.ag_id=u.ag_id) WHEN MATCHED THEN UPDATE SET j.unit_count=u.unit_count";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The final Console.WriteLine is for debugging purposes and is commented out of production and test code
        // Console.WriteLine("Step 6 exiting at: " + DateTime.Now.ToString("G");
      }
    }

    // Once the calls are inserted or merged into the evaluation table, this function purges the curent_temp table for reuse. This should help prevent data contamination
    // by ensuring stable sets of rows can be consistently obtained with each pass.
    private static void PurgeCurrentCalls()
    {
      // The Console.WriteLine function is included for debugging purposes and is not run in production or test code. 
      // Console.WriteLine("Step 7 entering at " + DateTime.Now.ToString("G");
      qry = "TRUNCATE TABLE hc_curent_temp REUSE STORAGE";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The final Console.WriteLine is for debugging purposes and is commented out of production and test code
        // Console.WriteLine("Step 7 exiting at: " + DateTime.Now.ToString("G");
      }
    }

    // This function purges out the contents of the hc_comment_temp table. This keeps memory demands lower and allows for quicker
    // gathering and comparison of data.
    private static void PurgeCurrentComments()
    {
      // The Console.WriteLIne function is included for debugging purposes and is not run in production or test code.
      // Console.WriteLine("Step 8 entering at " + DateTime.Now.ToString("G");
      qry = "TRUNCATE TABLE hc_comment_temp REUSE STORAGE";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The final Console.WriteLine is for debugging purposes and is commented out of production and test code.
        // Console.WriteLine("Step 8 exiting at: " + DateTime.Now.ToString("G");
      }
    }

    // This function purgest out the contents of the hc_unitcount_temp table. This keeps memory demands lower and allows for quicker
    // gathering and comparison of data.
    private static void PurgeCurrentCount()
    {
      // The Console.WriteLine function is included for debugging purposes and is not run in production or test code.
      // Console.WriteLine("Step 9 entering at: " + DateTime.Now.ToString("G");
      qry = "TRUNCATE TABLE hc_unitcount_temp REUSE STORAGE";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The final Console.WriteLine is for debugging purposes and is commented out of production and test code.
        // Console.WriteLine("Step 9 exiting at: " + DateTime.Now.ToString("G");
      }
    }

    // This function deletes completed calls (calls where the closed datetime stamp (xdts) is not empty) so they will not be re-examined.
    private static void CleanHotCalls()
    {
      // The Console.WriteLine function is included for debugging purposes and is not run in production or test code.
      // Console.WriteLine("Step 10 entering at: " + DateTime.Now.ToString("G");
      qry = "DELETE FROM jc_hc_curent WHERE xdts IS NOT NULL";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The final Console.WriteLine is for debugging purposes and is commented out of production and test code.
        // Console.WriteLIne("Step 10 exiting at: " + DateTime.Now.ToString("G");
      }
    }

    // This function deletes all open call over two hours old from the jc_hc_curent table so they will not be re-examined.
    private static void DeleteHotCalls()
    {
      // The Console.WriteLine function is included for debugging purposes and is not run in production or test code.
      // Console.WriteLine("Step 11 entering at: " + DateTime.Now.ToString("G");
      qry = "DELETE FROM jc_hc_curent WHERE substr(ad_ts,1,14) < TO_CHAR(systimestamp-2/24, 'YYMMDDHH24MISS')";
      OracleCommand cmd = new OracleCommand(qry, cxn);
      cxn.Open();
      OracleTransaction txn = cxn.BeginTransaction(IsolationLevel.ReadCommitted);
      try
      {
        cmd.ExecuteNonQuery();
        txn.Commit();
      }
      catch (OracleException ox)
      {
        MailException(ox);
        LogException(ox);
        txn.Rollback();
      }
      catch (Exception ex)
      {
        MailException(ex);
        LogException(ex);
        txn.Rollback();
      }
      finally
      {
        txn.Dispose();
        cmd.Dispose();
        cxn.Close();
        cxn.Dispose();
        // The final Console.WriteLine is for debugging purposes and is commented out of production and test code.
        // Console.WriteLine("Step 11 exiting at: " + DateTime.Now.ToString("G");
      }
    }

    // The following sections are designed to move the Mailing and Logging features out of the individual methods and make them reusable.
    public static void MailException(OracleException ox)
    {
      MailMessage oMsg = new MailMessage();
      oMsg.From = new MailAddress("jcso.exception@jocogov.org", "Oracle Exception");
      oMsg.To.Add(new MailAddress("tony.dunsworht@jocogov.org", "Tony Dunsworth"));
      oMsg.To.Add(new MailAddress("carter.wetherington@jocogov.org", "Carter Wetherington"));
      oMsg.Subject = "Oracle Poller Exception";
      // Altered the original DateTime string to a standard .NET string for stability
      oMsg.Body = "Oracle has reported the following exception: " + ox.Number + " with the following explanation: " + ox.ToString() + " at " + DateTime.Now.ToString("G");
      SmtpClient oPost = new SmtpClient();
      oPost.Send(oMsg);
    }

    public static void MailException(Exception ex)
    {
      MailMessage eMsg = new MailMessage();
      eMsg.From = new MailAddress("jcso.exception@jocogov.org", "Oracle Exception");
      eMsg.To.Add(new MailAddress("tony.dunsworht@jocogov.org", "Tony Dunsworth"));
      eMsg.To.Add(new MailAddress("carter.wetherington@jocogov.org", "Carter Wetherington"));
      eMsg.Subject = "Hot Calls Poller Exception";
      // Altered the original DateTime string to a standard .NET string for stability
      eMsg.Body = "The Hot Calls Poller threw the following exception: " + ex.ToString() + " at " + DateTime.Now.ToString("G");
      SmtpClient ePost = new SmtpClient();
      ePost.Send(eMsg);
    }

    public static void LogException(OracleException ox)
    {
      FileStream pollErrorLog = null;
      pollErrorLog = File.Open(@"C:\Temp\pollErrorLog.txt", FileMode.Append, FileAccess.Write);
      StreamWriter pollErrorWrite = new StreamWriter(pollErrorLog);
      pollErrorWrite.WriteLine("Oracle has reported the following exception: " + ox.Number + " with the following explanation: " + ox.ToString() + " at " + DateTime.Now.ToString("G"));
      pollErrorWrite.Close();
      pollErrorLog.Close();
    }

    public static void LogException(Exception ex)
    {
      FileStream pollErrorLog = null;
      pollErrorLog = File.Open(@"C:\Temp\pollErrorLog.txt", FileMode.Append, FileAccess.Write);
      StreamWriter pollErrorWrite = new StreamWriter(pollErrorLog);
      pollErrorWrite.WriteLine("The Hot Calls Poller threw the following exception: " + ex.ToString() + " at " + DateTime.Now.ToString("G"));
      pollErrorWrite.Close();
      pollErrorLog.Close();
    }
  }
  // Change 20130211-1: Updated regular expression for 10-NN codes to read for asterisks behind the instance.
  // Change 20130211-2: Changed timing on unit count from 20 to 25 minutes.
  // Change 20130508-1: Added new regular expression for "**Event held for" removal from comments.
  // Change 20140815-1: Moved the Mailing and Log functions into separate functions for better reuse.
  // Change 20140815-2: Moved the string for the query and the connection both out to a pulic call in order to make things more streamlined. 
  // Overall result was the removal of approximately 46% of the original code base and a more streamlined efficient structure.
}