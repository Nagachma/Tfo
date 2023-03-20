import pandas as pd
import time
import traceback
import json
import datetime
import os, sys
import math  
from statistics import median
import pandas as pd


#file imports
from services.TFO_slides_graphs import tfo_config
from utils.Logger import get_logger

tfo_utils_logger = get_logger("tfo_utils_logger")
tfo_utils_logger.debug("Initialized the tfo_utils_logger creation")

def calculate_cpp(sum_value,number_of_patients):
    try:
        result=(sum_value/number_of_patients)/tfo_config.cpp_divide_value
        tfo_utils_logger.debug(result)
    except Exception as e:
        tfo_utils_logger.debug((str(traceback.format_exc())))
        result=0
    return result


def read_inputfile(input_filepath):
    try:
        if input_filepath.endswith(".xlsx") or input_filepath.endswith(".xls"):
            result=pd.read_excel(input_filepath)
        if input_filepath.endswith(".csv"):
            result=pd.read_csv(input_filepath)
    except Exception as e:
        tfo_utils_logger.debug((str(traceback.format_exc())))
        result=pd.DataFrame()
    return result