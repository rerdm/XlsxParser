import string

class DatePrefix:
    """
    Class will create a prefix as String
    Usage from other program:

    Import module of class of dataPrefix.py:
    - from datePrefix import DatePrefix

    datePfx = DatePrefix().getPrefix()

    :return (str)
    _2022-02-12
    """

    def __init__(self):

        try:
            import sys
        except:
            sys.exit("Sys module con not be imported")

        try:
            import datetime
        except:
            sys.exit("datetime module con not be imported")

        self.__full_date_time_string = str(datetime.datetime.now())
        self.__date_time_string = self.__full_date_time_string.split()[0]

    def get_prefix(self):

        date_time_string = self.__full_date_time_string.split()[0]

        return str(date_time_string +"_")
