import pandas as pd
import time
import datetime
import os, sys
import math  
from statistics import median
from lxml import etree
from pptx import Presentation
from pptx import Presentation
from pptx.chart.data import CategoryChartData,ChartData,BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE , XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.dml.color import ColorFormat, RGBColor

absolute_path = os.path.abspath(__file__ + "/../../../")
sys.path.append(absolute_path)
#file imports
from services.TFO_slides_graphs import tfo_config
from services.TFO_slides_graphs.tfo_utils import  read_inputfile
from utils.Logger import get_logger

tfo_regional_allocation_logger = get_logger("tfo_regional_allocation_logger")
tfo_regional_allocation_logger.debug("Initialized the tfo_regional_allocation_logger creation")

def filter_regional_allocationdata(input_filepath,slide_name):
    '''
    Process input file to get chart data.
    '''
    input_df=pd.read_excel(input_filepath,header=1,index_col=0,sheet_name=tfo_config.slide10_sheetname)
    input_df['index_col'] = input_df.index
    input_df= input_df.drop_duplicates()
    chartdata_dict=dict()
    for category_name, category_rows in tfo_config.slide10_category_columnmap.items():
        plannedsite_values=[ input_df.loc[category_row][tfo_config.allocation_columnname]  for category_row in category_rows  if  category_row in input_df.index ]
        sum_value=sum(plannedsite_values)
        tfo_regional_allocation_logger.debug(sum_value)
        series_value=sum_value/input_df.loc[tfo_config.slide10_total][tfo_config.allocation_columnname]
        tfo_regional_allocation_logger.debug(series_value)
        if series_value==0.0:
            continue
        chartdata_dict[category_name]=series_value
    tfo_regional_allocation_logger.debug(chartdata_dict)
    return chartdata_dict

def filter_tforegional_allocationdata(input_filepath,slide_name):
    '''
    Process input file to get chart data.
    '''
    input_df=pd.read_excel(input_filepath,header=1,index_col=0,sheet_name=tfo_config.ra_tfosheetname)
    input_df['index_col'] = input_df.index
    input_df= input_df.drop_duplicates()
    total_value=input_df.loc[tfo_config.ra_total][tfo_config.ra_allocation_columnname]
    chartdata_dict=dict()
    result_df=input_df.groupby([tfo_config.ra_categorycolumn]).sum()
    result_df[tfo_config.ra_categorycolumn]=result_df.index
    tfo_regional_allocation_logger.debug(result_df)
    for index, region_row in result_df.iterrows():
        sum_value=region_row[tfo_config.ra_allocation_columnname]
        tfo_regional_allocation_logger.debug(sum_value)
        series_value=sum_value/total_value
        if series_value==0.0:
            continue
        category_name=region_row[tfo_config.ra_categorycolumn]
        chartdata_dict[category_name]=series_value
    tfo_regional_allocation_logger.debug(chartdata_dict)
    return chartdata_dict

def create_piechartdata(chartdata_dict):
    '''
    Extracts data from input data dict and prepares chart data.
    '''
    # define chart data 
    chart_data = ChartData()
    series_list=[]
    category_list=[]
    for category, category_value in chartdata_dict.items():
        category_list.append(category)
        series_list.append(category_value)
    chart_data.categories=category_list
    chart_data.add_series(tfo_config.slide10_seriesname, tuple(series_list))    
    return chart_data

def create_regional_allocation_chartdata(input_filepath,chart_name,slide_name):
    '''
    Creates regional allocation chart  data based on input file and geneartes pie chart.
    '''
    if slide_name=="TfoRegionalAllocation":
        chartdata_dict=filter_tforegional_allocationdata(input_filepath,slide_name)
    else:
        chartdata_dict=filter_regional_allocationdata(input_filepath,slide_name)
    # define chart data 
    chart_data =create_piechartdata(chartdata_dict)
    #create chart
    series_number_dict=dict()
    
    return chart_data, series_number_dict

