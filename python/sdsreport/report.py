import os
import sys
import glob
import re
from datetime import datetime, date, time

args = sys.argv
log_list = glob.glob(args[1])

report_dic = {}
ftd_dic = {}

# regular expression pattern
# ThreadID, Time, ItemID, FileName, FileSize
regex_pattern_request_received = re.compile(r"(?P<ThreadID>\d+-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}).* \| Sanitization Request Received: Request ID: (?P<ItemID>[^,]+), Source: .*\\(?P<FileName>[^,\\]+), Size: (?P<FileSize>\d+).*")
# ThreadID, Time, ItemID, FileName
regex_pattern_sanitization_started = re.compile(r"(?P<ThreadID>\d+-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}).* \| Sanitization Started: Item ID: (?P<ItemID>[^,]+), Filename: (?P<FileName>[^,]+), .*")
# ThreadID, Time, FileName
regex_pattern_sanitization_done = re.compile(r"(?P<ThreadID>\d+-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}).*\| \[10020110\] Sanitization Done \(File (?P<FileName>.*?) sanitization process successfully ended\.\)")
# ThreadID, Time, ItemID, PublishFileName, FileType
regex_pattern_publish_done = re.compile(r"(?P<ThreadID>\d+-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}) \| 2 Info \| Publish Done: Items: \{ Item ID: (?P<ItemID>[^,]+), Filename: (?P<FileName>[^,]+), Type: (?P<FileType>.*) \}")
# ThreadID, Time, IncludedFile, Version
regex_pattern_ftd = re.compile(r"(?P<ThreadID>\d+\-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}) \|.*FTD result for .*\\(?P<IncludedFile>[^\\]+) is \#Library version: (?P<Version>.*)")
# ThreadID, Time, ItemID
regex_pattern_status_lt_74 = re.compile(r"(?P<ThreadID>\d+\-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}).* \| (?P<ItemID>[-0-9a-z]+)'s Status = (?P<Status>Done|Blocked)")
regex_pattern_status_ge_74 = re.compile(r"(?P<ThreadID>\d+\-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}).* \| GetStatus was called for ID:(?P<ItemID>[-0-9a-z]+)\. Status:(?P<Status>Done|Blocked)")
# ThreadID, Time, ItemID, PublishFileName, FileType, Reason, Details
regex_pattern_block_reason = re.compile(r"(?P<ThreadID>\d+-\d+) \| (?P<Time>\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}\.\d{3}).* \| Item Blocked\. Item ID: (?P<ItemID>[^,]+), Filename: (?P<FileName>.*?), Type: (?P<FileType>.*?), Reason: (?P<Reason>[^,]+), Details: (?P<Details>.*)$")

# version check
try:
    for log in log_list:
        with open(log, "r", encoding="utf_8_sig") as f:
            for line in f:
                m = regex_pattern_ftd.match(line)
                if m:
                    m.group('Version')
                    version = m.group('Version')
                    break
except:
    pass

if float(version[0:3]) >= 7.4:
    regex_pattern_status = regex_pattern_status_ge_74
else:
    regex_pattern_status = regex_pattern_status_lt_74


log_pattern_dic = {
    regex_pattern_request_received:
    lambda m: ("SanitizationLog", m.group('ItemID'),
               {"RequestReceivedTime":datetime.strptime(m.group('Time'), "%d/%m/%Y %H:%M:%S.%f"),
                "FileName":m.group('FileName'),
                "FileSize":m.group('FileSize')
                }),
    regex_pattern_sanitization_started:
    lambda m: ("SanitizationLog", m.group('ItemID'),
               {"ThreadID":m.group('ThreadID'),
                "SanitizationStartedTime":datetime.strptime(m.group('Time'), "%d/%m/%Y %H:%M:%S.%f"),
                "FileName":m.group('FileName')}),
    regex_pattern_publish_done:
    lambda m: ("SanitizationLog", m.group('ItemID'),
               {"PublishDoneTime":datetime.strptime(m.group('Time'), "%d/%m/%Y %H:%M:%S.%f"),
                "PublishFileName":m.group('FileName'),
                "FileType":m.group('FileType')}),
    regex_pattern_ftd:
    lambda m: ("FtdLog", m.group('ThreadID'),
               {"FtdTime":datetime.strptime(m.group('Time'), "%d/%m/%Y %H:%M:%S.%f"),
                "IncludedFile":m.group('IncludedFile'),
                "Version":m.group('Version')}),
    regex_pattern_status:
    lambda m: ("SanitizationLog", m.group('ItemID'),
               {"ResponseDoneTime":datetime.strptime(m.group('Time'), "%d/%m/%Y %H:%M:%S.%f"),
                "Status":m.group('Status')}),
    regex_pattern_block_reason:
    lambda m: ("SanitizationLog", m.group('ItemID'),
               {"BlockReason":"{0}|{1}".format(m.group('Reason'), m.group('Details').strip('[]'))})
}


def make_record(regex):
    """
    make each report unit dictionary
    :param regex: compiled regular expression pattern
    """
    flag, record_id, value_dic = "", "", {}
    m = regex.match(line)
    if m:
        flag, record_id, value_dic = log_pattern_dic[regex](m)
        if flag == "SanitizationLog":
            if record_id not in report_dic:
                report_dic[record_id] = {"ThreadID":"", "FileName":"", "FileSize":"", "FileType":"",
                                         "RequestReceivedTime":"", "SanitizationStartedTime":"",
                                         "SanitizationDoneTime":"", "PublishDoneTime":"",
                                         "ResponseDoneTime":"", "TotalSanitizingTime":"",
                                         "PublishFileName":"", "Status":"", "IncludedFiles":[],
                                         "BlockReason":""}
            report_dic[record_id].update(value_dic)
        elif flag == "FtdLog":
            if record_id not in ftd_dic:
                ftd_dic[record_id] = []
            ftd_dic[record_id].append(value_dic)


# read log and make record
for log in log_list:
    print("Starting to process {} ...".format(log), file=sys.stderr)
    try:
        with open(log, "r", encoding="utf_8_sig") as f:
            for line in f:
                for regex in log_pattern_dic:
                    make_record(regex)
    except:
        pass


for item_id in report_dic:
    process_start_time, Process_end_time = None, None
    thread_id = report_dic[item_id]["ThreadID"]

    # get process start/end time
    process_start_time = report_dic[item_id]["SanitizationStartedTime"]
    process_end_time = report_dic[item_id]["SanitizationDoneTime"] \
                           if report_dic[item_id]["SanitizationDoneTime"] \
                           else report_dic[item_id]["PublishDoneTime"]

    # calcurate PublishProcessSeconds
    if process_end_time and process_start_time:
        report_dic[item_id]["TotalSanitizingTime"] = process_end_time - process_start_time

    # count included files
    if process_start_time and process_end_time:
        if thread_id in ftd_dic:
            for ftd_entry in ftd_dic[thread_id]:
                if process_start_time <= ftd_entry["FtdTime"] and ftd_entry["FtdTime"] <= process_end_time:
                    report_dic[item_id]["IncludedFiles"].append(ftd_entry["IncludedFile"])


# output
print("ItemID", "FileName", "FileSize", "FileType",
      "RequestReceivedTime", "SanitizationStartedTime",
      "PublishDoneTime", "ResponseDoneTime", "TotalSanitizingTime",
      "PublishFileName", "IncludedFileCount","Status", "BlockReason", sep=',')
for item_id in report_dic:
    print(item_id,
          report_dic[item_id]["FileName"].encode('utf_8_sig', 'ignore').decode('utf_8_sig', 'ignore'),
          report_dic[item_id]["FileSize"],
          report_dic[item_id]["FileType"],
          "" if not report_dic[item_id]["RequestReceivedTime"] else \
              datetime.strftime(report_dic[item_id]["RequestReceivedTime"], "%Y/%m/%d %H:%M:%S.%f").rstrip("000"),
          "" if not report_dic[item_id]["SanitizationStartedTime"] else \
              datetime.strftime(report_dic[item_id]["SanitizationStartedTime"], "%Y/%m/%d %H:%M:%S.%f").rstrip("000"),
          "" if not report_dic[item_id]["PublishDoneTime"] else \
              datetime.strftime(report_dic[item_id]["PublishDoneTime"], "%Y/%m/%d %H:%M:%S.%f").rstrip("000"),
          "" if not report_dic[item_id]["ResponseDoneTime"] else \
              datetime.strftime(report_dic[item_id]["ResponseDoneTime"], "%Y/%m/%d %H:%M:%S.%f").rstrip("000"),
          report_dic[item_id]["TotalSanitizingTime"],
          # report_dic[item_id]["IncludedFiles"].encode('utf_8_sig', 'ignore').decode('utf_8_sig', 'ignore'),
          report_dic[item_id]["PublishFileName"].encode('utf_8_sig', 'ignore').decode('utf_8_sig', 'ignore'),
          "" if not report_dic[item_id]["IncludedFiles"] else \
              len(report_dic[item_id]["IncludedFiles"]) - 1,
          report_dic[item_id]["Status"],
          report_dic[item_id]["BlockReason"],
          sep=','
          )

print("Completed.", file=sys.stderr)
