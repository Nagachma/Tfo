import pandas as pd
import time
import datetime
import os, sys
import math
import copy  
from statistics import median , mean
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
from services.TFO_slides_graphs.tfo_utils import calculate_cpp , read_inputfile
from utils.Logger import get_logger
tfo_cppsample_logger = get_logger("tfo_cppsample_logger")
tfo_cppsample_logger.debug("Initialized the tfo_cppsample_logger creation")



def filter_cppsample_data(input_filepath, input_json,planned_studycode_json):
    #constant variables
    columns_G_AB=tfo_config.columns_G_AB 
    study_status=tfo_config.study_status

    charts_data_dict=dict()
    yaxis_list=[]
    for status in study_status:
        input_df=read_inputfile(input_filepath)
        input_df.drop_duplicates(inplace=True)
        
        if status=="Planned –Trial scenario":
            data_pointlist=[]
            if planned_studycode_json["NoOfPatients"]:
                xaxis_value=int(planned_studycode_json["NoOfPatients"])
                yaxis_value=float(planned_studycode_json["CPP"])
                yaxis_list.append(yaxis_value)
                data_pointlist.append([xaxis_value,yaxis_value])            
        else:
            filtered_df=copy.deepcopy(input_df[(input_df['STUDY_STATUS']==status)])   
            filtered_df=filtered_df[(filtered_df["DEVELOPMENT_UNIT"]==input_json["DEVELOPMENT_UNIT"])
                    & (filtered_df["PROVIDING_ORGANIZATION"]==input_json["PROVIDING_ORGANIZATION"])
                    & (filtered_df["INDICATION"]==input_json["INDICATION"]) 
                    & (filtered_df["PHASE_DERIVED"]==input_json["PHASE_DERIVED"])
                    ]

            if filtered_df.empty:
                filtered_df=copy.deepcopy(input_df[(input_df['STUDY_STATUS']==status)])
                filtered_df=filtered_df[(filtered_df["DEVELOPMENT_UNIT"]==input_json["DEVELOPMENT_UNIT"])
                & (filtered_df["PROVIDING_ORGANIZATION"]==input_json["PROVIDING_ORGANIZATION"])
                & (filtered_df["PHASE_DERIVED"]==input_json["PHASE_DERIVED"])
                ]
            tfo_cppsample_logger.debug("filetered_df-----")
            tfo_cppsample_logger.debug(filtered_df)
            
            data_pointlist=[]
            
            if not filtered_df.empty:
                for index, excel_row in filtered_df.iterrows():                
                    number_of_patients=int(excel_row["PATIENTS"])
                    excel_row=excel_row[columns_G_AB]
                    tfo_cppsample_logger.debug("columns G to AB values----------")
                    tfo_cppsample_logger.debug(excel_row.values)
                    values=excel_row.values
                    sum_value=sum(value  for value in values if not (isinstance(value, str) or value!=value) )
                    yaxis_value=calculate_cpp(sum_value,number_of_patients)
                    xaxis_value=number_of_patients
                    yaxis_list.append(yaxis_value)
                    data_pointlist.append([xaxis_value,yaxis_value])
        charts_data_dict[status]=data_pointlist
    try:   
        median_value=median(yaxis_list)
    except:
        median_value=0.00
    try:
        average_value=mean(yaxis_list)
    except:
        average_value=0.00
    return charts_data_dict,median_value,yaxis_list,average_value


def filter_duplicaterows(points_list):
    '''
    filters datapoints -multiple entries for
    study code are combined if providing organization
    is different else any one row is considered.
    '''
    filtereddata_points=dict()
    for value in points_list:
        study_code=value[2]
        providing_organization=value[3]
        if study_code not in filtereddata_points.keys():
            filtereddata_points[study_code]=value
        elif study_code in filtereddata_points.keys():
            if filtereddata_points[study_code][3]==providing_organization:
                continue
            elif filtereddata_points[study_code][3]!=providing_organization:
                xaxis_value=(value[0]+filtereddata_points[study_code][0])/2
                yaxis_value=(value[1]+filtereddata_points[study_code][1])
                filtereddata_points[study_code][0]=xaxis_value
                filtereddata_points[study_code][1]=yaxis_value
    return filtereddata_points

def oncology_filter_cppsample_data(input_filepath, input_json,planned_studycode_json):
    #constant variables
    columns_G_AB=tfo_config.columns_G_AB 
    study_status=tfo_config.study_status

    charts_data_dict=dict()
    yaxis_list=[]
    for status in study_status:
        input_df=read_inputfile(input_filepath)
        input_df.drop_duplicates(inplace=True)
        
        if status=="Planned –Trial scenario":
            data_pointlist=[]
            if planned_studycode_json["NoOfPatients"]:
                xaxis_value=int(planned_studycode_json["NoOfPatients"])
                yaxis_value=float(planned_studycode_json["CPP"])
                yaxis_list.append(yaxis_value)
                data_pointlist.append([xaxis_value,yaxis_value])            
        else:
            filtered_df=copy.deepcopy(input_df[(input_df['STUDY_STATUS']==status)])
            for column_name in tfo_config.cppsample_oncologyfilters:    
                filtered_df=filtered_df[filtered_df[column_name].isin(input_json[column_name])]

            if filtered_df.empty:
                filtered_df=copy.deepcopy(input_df[(input_df['STUDY_STATUS']==status)])
                for column_name in tfo_config.cppsample_oncologyfilters:
                    if column_name in tfo_config.cppsample_ignorefilters:
                        continue    
                    filtered_df=filtered_df[filtered_df[column_name].isin(input_json[column_name])]

            tfo_cppsample_logger.debug("filetered_df-----")
            tfo_cppsample_logger.debug(filtered_df)
            
            data_pointlist=[]
            points_list=[]
            
            if not filtered_df.empty:
                for index, excel_row in filtered_df.iterrows():
                    study_code=excel_row["STUDY_CODE"]
                    providing_organization=excel_row["PROVIDING_ORGANIZATION"]                
                    number_of_patients=int(excel_row["PATIENTS"])
                    excel_row=excel_row[columns_G_AB]
                    tfo_cppsample_logger.debug("columns G to AB values----------")
                    tfo_cppsample_logger.debug(excel_row.values)
                    values=excel_row.values
                    sum_value=sum(value  for value in values if not (isinstance(value, str) or value!=value) )
                    yaxis_value=calculate_cpp(sum_value,number_of_patients)
                    xaxis_value=number_of_patients
                    yaxis_list.append(yaxis_value)
                    points_list.append([xaxis_value,yaxis_value,study_code,providing_organization])
                #filter values
                filtereddata_points=filter_duplicaterows(points_list)
                data_pointlist=[[value[0],value[1]] for key , value in filtereddata_points.items()]

        charts_data_dict[status]=data_pointlist
    try:   
        median_value=median(yaxis_list)
    except:
        median_value=0.00
    try:
        average_value=mean(yaxis_list)
    except:
        average_value=0.00
    return charts_data_dict,median_value,yaxis_list,average_value



def create_bubblechart_data(charts_data_dict):
    '''
    Extracts data from filtered data  and prepares chart data.
    '''
    bubble_weights=tfo_config.bubble_weights
    # define chart data 
    chart_data = BubbleChartData()
    series_number_dict=dict()
    index=-1
    for series_name, series_list in charts_data_dict.items():
        if series_list:
            index=index+1
            series_number_dict[index]=series_name
            series=chart_data.add_series(series_name)
            for data_point in series_list:
                series.add_data_point(data_point[0],data_point[1],bubble_weights[series_name])
    print(series_number_dict)
    return chart_data ,series_number_dict

def create_insert_cppsamplechart(input_filepath, input_json,  chart_name,planned_studycode_json):
    '''
    Creates cpp sample data based on input json and geneartes bubble chart.
    '''
    if input_json["SlideName"]=="CPP_ONCOLOGY":
        charts_data_dict,median_value,yaxis_list,average_value=oncology_filter_cppsample_data(input_filepath, input_json,planned_studycode_json)
    else:
        charts_data_dict,median_value,yaxis_list,average_value=filter_cppsample_data(input_filepath, input_json,planned_studycode_json)
    # define chart data 
    chart_data ,series_number_dict=create_bubblechart_data(charts_data_dict)

    return chart_data ,series_number_dict,median_value,yaxis_list ,average_value




