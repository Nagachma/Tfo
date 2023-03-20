import win32com.client
from win32com.client import constants
import pythoncom
import traceback
import pandas as pd
#file imports
from services.TFO_slides_graphs import tfo_config
from utils.Logger import get_logger

tfo_embedexcel_logger = get_logger("tfo_embedexcel_logger")
tfo_embedexcel_logger.debug("Initialized the tfo_embedexcel_logger creation")

def get_excelrange(excel_filepath,sheet_name):
    input_df=pd.read_excel(excel_filepath,sheet_name=sheet_name,header=None)
    input_df.columns=tfo_config.static_columnnames_list[0:len(input_df.columns)]
    
    data_range=None
    if sheet_name=="Country_Allocation":
        last_columnname=input_df.columns[-1]
        end_range=last_columnname+str(len(input_df))
        data_range=tfo_config.data_range.format(end_range)
        
    elif sheet_name=="Cost_Summary":
        #To remove comments column
        last_columnname=input_df.columns[-2]
        tfo_embedexcel_logger.debug(last_columnname)        
        column_name=tfo_config.cost_summary_columnname
        input_df[column_name]=input_df[column_name].str.strip()
        index_list=input_df.index[input_df[column_name] == tfo_config.first_check].tolist()
        if not index_list:
            index_list=input_df.index[input_df[column_name] == tfo_config.second_check].tolist()
            if not index_list:
                return None
            else:
                end_range=last_columnname+str(index_list[0]+1)
        else:
            end_range=last_columnname+str(index_list[0]-1)
        data_range=tfo_config.data_range.format(end_range)
        tfo_embedexcel_logger.debug(data_range)
    return data_range

def embed_excel_inpptx(sheet_name,excel_filepath,pptx_filepath):
    data_range=None
    try:
        data_range=get_excelrange(excel_filepath,sheet_name)
        if data_range is None:
            return pptx_filepath

        pythoncom.CoInitialize()
        powerpoint_object = win32com.client.Dispatch("Powerpoint.Application")
        #read and create pptx file object.
        powerpoint_presentation = powerpoint_object.Presentations.Open(pptx_filepath,WithWindow=False)
        #read abd create excelfile object.
        excel_object = win32com.client.Dispatch("Excel.Application")
        excel_object.DisplayAlerts = False
        powerpoint_object.DisplayAlerts=False
        excel_workbook = excel_object.Workbooks.Open(Filename=excel_filepath)
        sheet_name=sheet_name
        excel_worksheet = excel_workbook.Worksheets(sheet_name)
        excel_range = excel_worksheet.Range(data_range)
        excel_range.Copy()

        powerpoint_slide = powerpoint_presentation.Slides[0]
        powerpoint_slide.Shapes.PasteSpecial(DataType=10)
    except Exception as e:
        tfo_embedexcel_logger.debug(str(traceback.format_exc()))
    finally:
        if data_range is not None:
            powerpoint_presentation.Save()
            powerpoint_presentation.Close()
            excel_workbook.Close()
            pythoncom.CoUninitialize()
        
    return pptx_filepath





