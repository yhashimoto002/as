import os
import sys
import glob
import re
import csv
from datetime import datetime, date, time

# args
args = sys.argv
if len(args) == 1 or len(args) > 3:
    print("Invalid arguments!")
    print(r"Usage: .\collect_image_size_of_pdf.py *log*")
    print(r"Usage: .\collect_image_size_of_pdf.py *log* report.csv")
    exit(1)

if len(args) == 3:
    log_name = args[2]
else:
    log_name = 'report.csv'

log_list = glob.glob(args[1])

# variables
report_dic = {}
sanitization_done_dic = {}
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
    regex_pattern_sanitization_done:
    lambda m: ("SanitizationDoneLog", m.group('ThreadID'),
               {"SanitizationDoneTime":datetime.strptime(m.group('Time'), "%d/%m/%Y %H:%M:%S.%f"),
                "PublishFileName":m.group('FileName')}),
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
                                         "ResponseDoneTime":"", "TotalProcessSeconds":"",
                                         "UploadAndQueueWaitSeconds":"", "SanitizationProcessSeconds":"",
                                         "PublishProcessSeconds":"", "DownloadWaitSeconds":"",
                                         "PublishFileName":"", "Status":"", "IncludedFiles":[],
                                         "BlockReason":""}
            report_dic[record_id].update(value_dic)
        elif flag == "SanitizationDoneLog":
            if record_id not in sanitization_done_dic:
                sanitization_done_dic[record_id] = []
            sanitization_done_dic[record_id].append(value_dic)
        elif flag == "FtdLog":
            if record_id not in ftd_dic:
                ftd_dic[record_id] = []
            ftd_dic[record_id].append(value_dic)


# read log and make record
start_time = datetime.now()
try:
    for log in log_list:
        print("Starting to process {} ...".format(log), file=sys.stderr)
        with open(log, "r", encoding="utf_8_sig") as f:
            for line in f:
                for regex in log_pattern_dic:
                    make_record(regex)
except FileNotFoundError:
    pass


with open(log_name, 'w', encoding='utf_8_sig') as f:
    f.write("ItemID, FileName, FileSize, FileType, RequestReceivedTime, SanitizationStartedTime, SanitizationDoneTime, PublishDoneTime, ResponseDoneTime, TotalProcessSeconds, UploadAndQueueWaitSeconds, PublishProcessSeconds, DownloadWaitSeconds, PublishFileName, IncludedFileCount, Status, BlockReason\n")


print("Creating csv file ...")
for item_id in report_dic:
    thread_id = report_dic[item_id]["ThreadID"]

    # Calculationg TotalProcessSeconds.
    if report_dic[item_id]["RequestReceivedTime"] and report_dic[item_id]["ResponseDoneTime"]:
        report_dic[item_id]["TotalProcessSeconds"] = (report_dic[item_id]["ResponseDoneTime"] - report_dic[item_id]["RequestReceivedTime"]).total_seconds()
    
    # Calculationg UploadAndQueueWaitSeconds.
    if report_dic[item_id]["RequestReceivedTime"] and report_dic[item_id]["SanitizationStartedTime"]:
        report_dic[item_id]["UploadAndQueueWaitSeconds"] = (report_dic[item_id]["SanitizationStartedTime"] - report_dic[item_id]["RequestReceivedTime"]).total_seconds()
    
    # Calculating PublishProcessSeconds.
    if report_dic[item_id]["SanitizationStartedTime"] and report_dic[item_id]["PublishDoneTime"]:
        report_dic[item_id]["PublishProcessSeconds"] = (report_dic[item_id]["PublishDoneTime"] - report_dic[item_id]["SanitizationStartedTime"]).total_seconds()
    
    # Calculating DownloadWaitSeconds.
    if report_dic[item_id]["PublishDoneTime"] and report_dic[item_id]["ResponseDoneTime"]:
        report_dic[item_id]["DownloadWaitSeconds"] = (report_dic[item_id]["ResponseDoneTime"] - report_dic[item_id]["PublishDoneTime"]).total_seconds()

    # Calculating SanitizationDoneTime.
    if thread_id in sanitization_done_dic:
        for sanitization_done_entry in sanitization_done_dic[thread_id]:
            if report_dic[item_id]["SanitizationStartedTime"] <= sanitization_done_entry["SanitizationDoneTime"] <= report_dic[item_id]["PublishDoneTime"]:
                report_dic[item_id]["SanitizationDoneTime"] = sanitization_done_entry["SanitizationDoneTime"]

    # count included files
    if thread_id in ftd_dic:
        for ftd_entry in ftd_dic[thread_id]:
            if report_dic[item_id]["SanitizationStartedTime"] <= ftd_entry["FtdTime"] <= report_dic[item_id]["PublishDoneTime"]:
                #if ftd_entry["IncludedFile"] not in report_dic[item_id]["IncludedFiles"]:
                #    report_dic[item_id]["IncludedFiles"].append(ftd_entry["IncludedFile"])
                report_dic[item_id]["IncludedFiles"].append(ftd_entry["IncludedFile"])
    
    included_file_count = 0 if not report_dic[item_id]["IncludedFiles"] else len(report_dic[item_id]["IncludedFiles"]) - 1

    # output
    report_format = {"ItemID": item_id,
                     "FileName": report_dic[item_id]["FileName"],
                     "FileSize": report_dic[item_id]["FileSize"],
                     "FileType": report_dic[item_id]["FileType"],
                     "RequestReceivedTime": report_dic[item_id]["RequestReceivedTime"],
                     "SanitizationStartedTime": report_dic[item_id]["SanitizationStartedTime"],
                     "SanitizationDoneTime": report_dic[item_id]["SanitizationDoneTime"],
                     "PublishDoneTime": report_dic[item_id]["PublishDoneTime"],
                     "ResponseDoneTime": report_dic[item_id]["ResponseDoneTime"],
                     "TotalProcessSeconds": report_dic[item_id]["TotalProcessSeconds"],
                     "UploadAndQueueWaitSeconds": report_dic[item_id]["UploadAndQueueWaitSeconds"],
                     "PublishProcessSeconds": report_dic[item_id]["PublishProcessSeconds"],
                     "DownloadWaitSeconds": report_dic[item_id]["DownloadWaitSeconds"],
                     "PublishFileName": report_dic[item_id]["PublishFileName"],
                     "IncludedFileCount": included_file_count,
                     "Status": report_dic[item_id]["Status"],
                     "BlockReason": report_dic[item_id]["BlockReason"]}
    with open(log_name, 'a', encoding='utf_8_sig') as f:
        header = report_format.keys()
        writer = csv.DictWriter(f, fieldnames=header, lineterminator='\n')
        writer.writerow(report_format)


end_time = datetime.now()

# how many hours take
print("Start: {0}".format(start_time.strftime("%Y/%m/%d %H:%M:%S")))
print("End: {0}".format(end_time.strftime("%Y/%m/%d %H:%M:%S")))
print("Total: {0}".format(end_time - start_time))
