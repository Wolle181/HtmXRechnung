"""
Module WindowsEventLog
"""

import logging

try:
    import win32evtlog
    import win32evtlogutil
    HAS_WIN32_EVENTLOG = True
except ImportError:
    HAS_WIN32_EVENTLOG = False


class WindowsEventLogHandler(logging.Handler):
    def __init__(self, appname: str, logtype: str = "Application") -> None:
        logging.Handler.__init__(self)
        self.appname = appname
        self.logtype = logtype

    def emit(self, record) -> None:
        if not HAS_WIN32_EVENTLOG:
            return
        
        try:
            msg = self.format(record)
            eventID = record.levelno
            eventCategory = 0
            eventType = win32evtlog.EVENTLOG_INFORMATION_TYPE

            if record.levelno >= logging.ERROR:
                eventType = win32evtlog.EVENTLOG_ERROR_TYPE
            elif record.levelno >= logging.WARNING:
                eventType = win32evtlog.EVENTLOG_WARNING_TYPE

            win32evtlogutil.ReportEvent(
                self.appname, eventID, eventCategory, eventType, [msg]
            )
        except Exception:
            self.handleError(record)
