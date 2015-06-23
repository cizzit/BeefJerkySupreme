import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import time
import logging
import tempfile

import pyodbc
import smtplib
from tabulate import tabulate

logging.basicConfig(
    filename=tempfile.gettempdir()+'\ExceptionSvc.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)-7.7s %(message)s'
)


class DFSSVC(win32serviceutil.ServiceFramework):
    _svc_name_ = "ABBYYExceptionCheck"
    _svc_display_name_ = "ABBYY Exception Check"
    _svc_description_ = "Checks for batches in the exceptions queue and notifies support"
    database = {
        'server': 'localhost',
        'name': 'FlexiCapture11',
        'user': 'USERNAME',
        'pass': 'PASSWORD'
    }
    email_server = 'MAILSERVER'
    email_user = 'MAILUSER'
    email_pass = 'MAILPASS'
    email_recipients = ['LISTOFRECIPIENTS']
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(5)
        self.stop_requested = False
        
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STOPPED,
            (self._svc_name_, '')
        )
        self.stop_requested = True
        logging.warning('*** Service stop ***')
        
    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        self.main()
        
    def main(self):
        """
        Main service loop. Sleeps for 59 seconds out of 60 and runs a state DB check for any batches in exception
        status. If one is found, perform sanity check against list in memory and if it's new or has changed,
        update the list and send an email alert.
        If no batches in exception were found, empty the list in memory and re-loop until the service is stopped.
        :return: None
        """
        logging.warning('** Service start **')
        exceptions_list = []
        while not self.stop_requested:
            new_exception = False
            for i in range(0, 60):
                """
                I know this looks gumby but it means that the while loop isn't blocked by a time.sleep(60) when any
                stop message comes in.
                This way the while loop goes until it gets the signal, while the query only runs every 60 seconds
                """
                if not i % 60:
                    data = self.sql_exception_check()
                    if data:
                        _exception_flag = False
                        for d in data:
                            count, state, project = d
                            if 'Exceptions' in state:
                                _exception_flag = True
                                if not any(el['project'] == project for el in exceptions_list):
                                    # project doesn't exist in exceptions_list yet
                                    logging.info('Exceptions found, adding to list: %s, count %d' % (
                                        project,
                                        count
                                    ))
                                    exceptions_list.append(
                                        {
                                            'count': count,
                                            'state': state,
                                            'project': project
                                        }
                                    )
                                    new_exception = True
                                else:
                                    # project already exists in the list
                                    for el in exceptions_list:
                                        if project in el['project'] and count != el['count']:
                                            old_count, el['count'] = el['count'], count
                                            new_exception = True
                                            logging.info('Exception exists for %s, updating count (%d->%d)' % (
                                                project,
                                                old_count,
                                                count
                                            ))
                                        # else:
                                        #     logging.info('No change in exception list, skipping ...')
                        if exceptions_list and not _exception_flag:
                            logging.info('No exceptions reported, clearing memory')
                            exceptions_list = []

                if exceptions_list and new_exception:
                    self.send_mail(self.generate_table_data(exceptions_list))
                    new_exception = False
                time.sleep(1)
        return

    def send_mail(self, body_text):
        """
        Handles the sending of mail alerts
        :param body_text: string of text to incorporate the email body
        :return: None
        """
        message = (
            "From: ABBYY Monitor <no_reply@whatever.info>\r\n" +
            "To: " + ', '.join(self.email_recipients) + "\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/html\r\n" +
            "Subject: [ABBYY] !! Batches in Exception !!\r\n\r\n" +
            "<pre>Below is a table showing the batches currently in Exception state.\r\n" +
            body_text +
            "\r\nThis was generated at %s</pre>\r\n" % time.asctime()
        )
        try:
            smtphost = self.email_server
            smtpobj = smtplib.SMTP(smtphost)
            smtpobj.set_debuglevel(False)
            smtpobj.ehlo()
            if smtpobj.has_extn('STARTTLS'):
                smtpobj.starttls()
                smtpobj.ehlo()
            smtpobj.sendmail('no_reply@whatever.info', self.email_recipients, message)
            logging.info('Email alert sent.')
        except Exception as e:
            logging.error('Send mail error: %s' % e)

    @staticmethod
    def generate_table_data(data):
        """
        :param data: tuple containing state data
        :return: string
        Moved to function for future processing
        """
        return tabulate(data, headers="keys", tablefmt="psql")
        
    def sql_exception_check(self):
        """
        Query database for Batch Stages and return the result
        :return: list of tuples
        """
        query = "SELECT COUNT(b.Id) AS 'BatchCount', " \
                  "ps.Name AS 'Processing Stage', " \
                  "p.Name AS 'Project' " \
                  "FROM Batch as b " \
                  "JOIN ProcessingStage as ps on (b.ProcessingStageId = ps.Id) " \
                  "JOIN Project AS p on (b.ProjectId = p.Id) " \
                  "GROUP BY ps.Name, p.Name"
        try:
            con = pyodbc.connect(self.get_connection_string())
        except Exception as e:
            logging.error('Connection failure: %s' % e)
            return False
        try:
            cur = con.cursor()
            cur.execute(query,[])
            data = cur.fetchall()
        except Exception as e:
            logging.error('Query error: %s (Q: %s)' % (e, query))
            data = False
        finally:
            if con:
                con.close()
        return data

    def get_connection_string(self):
        """
        Build the database connection string. Seperated for readability
        :return: string
        """
        return r"DRIVER={SQL Server};" \
               r"SERVER=%(server)s;" \
               r"DATABASE=%(name)s;" \
               r"UID=%(user)s;" \
               r"PWD=%(pass)s" % self.database
        
if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(DFSSVC)

