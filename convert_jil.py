#!/usr/bin/env python3
# -*- coding:utf-8 -*-
'''
# File: jobmanager.py
# Project: AUTOJOB
# Created Date: 2022-09-27, 04:16:10
# Author: Chungman Kim(h2noda@unipark.kr)
# Last Modified: Fri Oct 21 2022
# Modified By: Chungman Kim
# Copyright (c) 2022 Unipark
# HISTORY:
# Date      	By	Comments
# ----------	---	----------------------------------------------------------
'''

from asyncio.log import logger

from Sloppy.message import *
from Sloppy.excel import *
from Sloppy.json import *
from Sloppy.common import *

import argparse
import os.path
import re

import logging

from openpyxl.utils import get_column_letter
from rich.progress import track

JOB_FIELD = "./jobfield.json"


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--JIL", "-J", help="JIL Filename")
    args = parser.parse_args()

    return args


def convert_filename(filename, filetype):
    """파일의 확장자를 변경

    Args:
        filename (String): 대상 파일명
        filetype (String): 바뀔 확장자명

    Returns:
        String: 변경된 파일명
    """
    name, ext = os.path.splitext(filename)
    retval = name + "." + filetype
    return retval


def initialize(args):

    global arg_jilfile
    global arg_excelfile
    global used_jobfield
    global dq_jobfield
    global xlsx
    global wb
    global ws

    arg_jilfile = args.JIL
    arg_excelfile = convert_filename(arg_jilfile, "xlsx")

    try:
        if os.path.isfile(arg_jilfile) == False:
            raise FileNotFoundError
        else:
            msg.printmsg("Jil File : " + arg_jilfile, "->", 1, "white", False)
            xlsx = Xlsx(arg_excelfile, "JOB", mode="n")
            wb = xlsx.wb
            ws = xlsx.ws

        used_jobfield = read_used_jobfield(JOB_FIELD)
        dq_jobfield = read_dq_jobfield(JOB_FIELD)
    except FileNotFoundError:
        msg.printmsg("Error : Convert File Not Found!!", "", 0, "red", True)
        exit()


def read_used_jobfield(jobfield):
    """Job 정보에서 사용될 Field를 읽음
        - "used" 값이 "Y" : Excel 파일에 기록

    Args:
        jobfield (String): 사용될 Job Field

    Returns:
        _type_: _description_
    """
    job_field = Json.read_jsonfile(jobfield)
    used_field = {}
    used_field = [i["field"] for i in job_field["field"] if i["used"] == "Y"]
    return used_field


def read_dq_jobfield(jobfield):
    """Double Quotation 사용 필드 설정

    Args:
        jobfield (String): Job Field

    Returns:
        _type_: _description_
    """
    job_field = Json.read_jsonfile(jobfield)
    dq_jobfield = {}
    dq_jobfield = [i["field"]
                   for i in job_field["field"] if (i["used"] == "Y" and i["double_quotation"] == "Y")]
    return dq_jobfield


def count_job(filename):
    """jil 파일 내의 Job 수량을 확인
        (insert_job 단어 기준)

    Args:
        filename (String): 파일명

    Returns:
        int: Job 수량
    """
    with open(filename, "r") as file:
        data = file.read()
        cnt_job = data.count("insert_job:")
    return cnt_job


def convert_j2e():
    ws.append(used_jobfield)

    cnt_job = count_job(arg_jilfile)
    list_alljob = []
    job = {}    # 개별 Job 정보
    with open(arg_jilfile, "r") as jil:
        for line_jil in track(jil, description="[white]  -> Convert Job Info :", total=cnt_job):
            if "insert_job:" in line_jil:
                list_alljob.append(job)
                line_jil = line_jil.strip()
                jobName = re.findall(r"insert_job:(.*?)job_type:", line_jil)[0]
                jobType = line_jil.split("job_type:")[1]
                job = {}
                job["insert_job"] = str(jobName).strip()
                job["job_type"] = str(jobType).strip()
            else:
                line_jil = line_jil.strip()
                if line_jil != "\n" and "/* ----" not in line_jil and line_jil != "":
                    if "start_times" in line_jil:
                        spli = line_jil.split("start_times:")
                        job["start_times"] = str(
                            spli[1]).replace("\"", "").strip()
                    elif "permission:" in line_jil:
                        spli = line_jil.split("permission:")
                        spli_value = str(spli[1]).strip()
                        if spli_value:
                            job["permission"] = str(spli[1]).strip()
                        else:
                            job["permission"] = ""
                    elif "command:" in line_jil:
                        spli = line_jil.split("command:")
                        job["command"] = str(spli[1]).strip()
                    # notification_emailaddress 처리
                    elif "notification_emailaddress" in line_jil:
                        spli = line_jil.split(
                            "notification_emailaddress:", 1)
                        if "notification_emailaddress" not in job:
                            job["notification_emailaddress"] = str(
                                spli[1]).strip()
                        else:
                            job["notification_emailaddress"] += ", " + \
                                str(spli[1]).strip()
                    else:
                        spli = line_jil.split(":", 1)
                        job[str(spli[0]).strip()] = str(
                            spli[1]).strip().replace("\"", "")

        list_alljob.append(job)

    list_alljob.pop(0)
    for ar in list_alljob:
        values = []
        for k in used_jobfield:
            if k in ar:
                values.append(ar[k])
            else:
                values.append(None)
        ws.append(values)

    msg.printmsg("Convert Job : " + str(cnt_job), "->", 1)
    msg.printmsg("Save Excel File", "->", 1, "white", False)

    # Resize - Excel File Column Width
    for col in ws.columns:
        max_length = 0
        column = col[0].column
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))

        column_letter = get_column_letter(column)
        adjusted_width = (max_length + 1) * 1.1

        if adjusted_width > 50:
            ws.column_dimensions[column_letter].width = 50
        else:
            ws.column_dimensions[column_letter].width = adjusted_width
    wb.save(arg_excelfile)


def main():
    global logger

    logger = logging.getLogger()

    try:
        args = get_args()
        # Step 00. Title
        msg.title("   Autosys Job Converter   ", color="blue")

        # Step 01. Initialize(Load Rule / Config, Env, etc)
        msg.printmsg("Step #01. Initialize...", "", 0, "green", True)

        msg.printmsg("Load Environment variable", "->", 1)
        initialize(args)

        # Step 02. Convert
        msg.printmsg("Step #02. Convert...", "", 0, "green", True)
        msg.printmsg("Convert Type : JIL to Excel", "->", 1, "white", False)
        convert_j2e()
        msg.printmsg("Finished....", "->", 1)

    except Exception as e:
        logger.exception(str(e))
        msg.printmsg("Error...Error...Error...", "", 0, "red", True)

    else:
        # Step 99. OK
        msg.blank()
        msg.OK()

    finally:
        msg.blank()
        msg.copyright()
        msg.blank()


if __name__ == "__main__":

    msg = Msg()
    main()
