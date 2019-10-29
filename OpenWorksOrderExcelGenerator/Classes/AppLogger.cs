using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;

using NLog;

namespace BomExcelGenerator
{
    public class AppLogger
    {
        static Logger logger = LogManager.GetLogger("appLogger");
        static Logger mailer = LogManager.GetLogger("appMailer");

        public static void ReportError(string message, bool email = false)
        {
            logger.Error(message);
        }

        public static void ReportWarning(string message, bool email = false)
        {
            logger.Warn(message);
        }

        public static void ReportInfo(string message, bool email = false)
        {
            logger.Info(message);
        }

        public static void SendNLogEmail(string message)
        {
            try
            {
                //NLog.Targets.MailTarget mailTarget = (NLog.Targets.MailTarget)FacadeLogger.mailer;
                string body = "<h1>This is a test email from NLog and Chris Weeks</h1>";
                body += "<br/>" + message;
                //NLog.Targets.MailTarget t = mailer.tar mailer.Factory.Configuration.AllTargets.First() as NLog.Targets.MailTarget;
                NLog.Targets.MailTarget t = ((NLog.Targets.MailTarget) LogManager.Configuration.FindTargetByName("appMail"));
                t.Subject = "Shortage Report Generation Nightly Report";
                mailer.Error(body);
            }
            catch (Exception ex)
            {
                logger.Error("Mail Error. Details : " + ex.Message);
            }
        }

        public static void SendShortageReportGenerationUpdateEmail(string message, string header, string prefix, string dataType, string verb, int successes = 0, int failures = 0)
        {
            try
            {
                string body = "<h2>" + header + "</h2>";
                body += message;
                if (successes != 0 || failures != 0)
                    body += "<p><ul><li>There were " + successes + " successful " + dataType + " " + verb + ".</li><li>There were " + failures + " failed " + dataType + " " + verb + ".</li></ul><p>Please see the error logs for further details at <a href='\\\\thas-connect01\\ShortageReports\\Logging\\Logs\\'>\\\\thas-connect01\\ShortageReports\\Logging\\Logs\\</a></p>";

                body += "<p><i>This is an automatically generated email from TAS Connect.<b>Please do not reply to this email directly.</b></i></p><p>Thanks,<br/>TAS Software Development Team</p>";
                NLog.Targets.MailTarget myTarget = ((NLog.Targets.MailTarget)LogManager.Configuration.FindTargetByName("appMail"));
                myTarget.Subject = "Shortage Report Generation Nightly Report - " + prefix;
                mailer.Error(body);
            }
            catch (Exception ex)
            {
                logger.Error("Mail Error. Details : " + ex.Message);
            }
        }
    }
}