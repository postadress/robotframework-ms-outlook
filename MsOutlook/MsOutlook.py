from robot.api.deco import keyword
import extract_msg
import re
from extract_msg.message import Message as OutlookMessage


class MsOutlook:
    ROBOT_LIBRARY_VERSION = '0.0.1'

    REGEX_BODY_PARSER = r"(?P<recipient_address>[\w\.-]*@[\w\.-]*\.[a-zA-Z]{2,3})[\r]?\n#.*smtp;(?P<smtp_response>\d{3}).*#SMTP#"
    BODY_PARSER = None

    def __init__(self):
        self.BODY_PARSER = re.compile(self.REGEX_BODY_PARSER)

    @keyword("Parse outlook message")
    def parse_message(self, file_path: str) -> OutlookMessage:
        """
        Returns a MSG object from an outlook msg file.
        """
        return extract_msg.Message(file_path)

    @keyword("Convert outlook msg to dictionairy")
    def convert_msgfile2dict(self, file_path: str, regex:str) -> dict:
        """
        Parses message body of a given outlook msg file. The regular expression must contain named groups.
        Return value is a dictionairy matching group names and values parsed from outlook message.
        """
        msg = self.parse_message(file_path)
        return self.convert_msg2dict(msg, regex)

    def convert_msg2dict(self, msg: OutlookMessage, regex: str)->dict:
        result = dict()

        parser = re.compile(regex)
        matches = parser.search(msg.body)

        if matches:
            for group in parser.groupindex.keys():
                result[group] = matches.group(group)
        return result
