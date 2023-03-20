import pandas as pd
import time
import datetime
import os, sys
import math
import copy
import traceback  
from statistics import median
from lxml import etree
from pptx import Presentation
from pptx import Presentation
from pptx.chart.data import CategoryChartData,ChartData,BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE , XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.dml.color import ColorFormat, RGBColor

#file imports
from services.TFO_slides_graphs import tfo_config
from services.TFO_slides_graphs.tfo_utils import calculate_cpp , read_inputfile
from utils.Logger import get_logger

tfo_cppsimilar_logger = get_logger("tfo_cppsimilar_logger")
tfo_cppsimilar_logger.debug("Initialized the tfo_cppsimilar_logger creation")

def filter_cppsimilartrails_data(input_filepath, input_json,planned_studycode, planned_cpp):
    '''
    Process input file to get chart data.
    '''
    columns_G_AB=tfo_config.columns_G_AB   
    input_df=read_inputfile(input_filepath)
    input_df.drop_duplicates(inplace=True)
    temp_dict=dict()
    study_code_list=[]
    yaxis_list=[]
    for study_code in input_json["STUDY_CODE"]:
        filtered_df=input_df[(input_df['STUDY_CODE']==study_code)]
        filtered_df=filtered_df[(filtered_df["DEVELOPMENT_UNIT"]==input_json["DEVELOPMENT_UNIT"])
                & (filtered_df["PROVIDING_ORGANIZATION"]==input_json["PROVIDING_ORGANIZATION"])
                & (filtered_df["INDICATION"]==input_json["INDICATION"]) 
                & (filtered_df["PHASE_DERIVED"]==input_json["PHASE_DERIVED"])
                ]
        if filtered_df.empty:
                filtered_df=input_df[(input_df['STUDY_CODE']==study_code)]
                filtered_df=filtered_df[(filtered_df["DEVELOPMENT_UNIT"]==input_json["DEVELOPMENT_UNIT"])
                & (filtered_df["PROVIDING_ORGANIZATION"]==input_json["PROVIDING_ORGANIZATION"])
                & (filtered_df["PHASE_DERIVED"]==input_json["PHASE_DERIVED"])
                ]

        excel_rows=filtered_df[filtered_df["STUDY_CODE"]==study_code]
        study_code_list.append(study_code)
        
        if not excel_rows.empty:
            #get first row
            excel_row=excel_rows.head(1)        
            number_of_patients=int(excel_row["PATIENTS"])
            excel_row=excel_row[columns_G_AB]
            tfo_cppsimilar_logger.debug("columns G to AB ------------------")
            tfo_cppsimilar_logger.debug(excel_row.iloc[0].values)
            values=excel_row.iloc[0].values            
            sum_value=sum(value  for value in values if not (isinstance(value, str) or value!=value) )
            yaxis_value=calculate_cpp(sum_value,number_of_patients)
            yaxis_list.append(yaxis_value)
        else:
            yaxis_list.append(0)
    #add primary study code to list
    study_code_list.append(planned_studycode)
    yaxis_list.append(float(planned_cpp))

    temp_dict["x_axis"]=study_code_list
    temp_dict["y_axis"]=yaxis_list
    median_value=median(yaxis_list)
    tfo_cppsimilar_logger.debug("median value---------------------")
    tfo_cppsimilar_logger.debug(median_value)
    processed_df=pd.DataFrame.from_dict(temp_dict)
    return processed_df


def filter_duplicaterows(points_list):
    '''
    filters datapoints -multiple entries for
    study code are combined if providing organization
    is different else any one row is considered.
    '''
    filtereddata_points=dict()
    for value in points_list:
        study_code=value[0]
        providing_organization=value[2]
        if study_code not in filtereddata_points.keys():
            filtereddata_points[study_code]=value
        elif study_code in filtereddata_points.keys():
            if filtereddata_points[study_code][2]==providing_organization:
                continue
            elif filtereddata_points[study_code][2]!=providing_organization:
                yaxis_value=(value[1]+filtereddata_points[study_code][1])
                filtereddata_points[study_code][1]=yaxis_value
    return filtereddata_points

def filterdf_values(input_df,filtered_df,input_json,studycode_flag=True):
    filtered_df=copy.deepcopy(input_df)
    if input_json['PROJECT_CODE']:
        filtered_df=filtered_df[filtered_df['PROJECT_CODE'].str.startswith(input_json['PROJECT_CODE'])]

    if studycode_flag:
        filtered_df=filtered_df[filtered_df['STUDY_CODE'].isin(input_json['STUDY_CODE'])]

    
    filterednew_df=copy.deepcopy(filtered_df)
    for column_name in tfo_config.cppsimilartrials_oncologyfilters:
        if input_json[column_name]:    
            filterednew_df=filterednew_df[filterednew_df[column_name].isin(input_json[column_name])]

    if filterednew_df.empty:
        filterednew_df=copy.deepcopy(filtered_df)
        #ignore indication filter
        for column_name in tfo_config.cppsimilartrials_oncologyfilters:
            if column_name in tfo_config.cppsimilartrails_ignorefilters:
                continue
            if input_json[column_name]:    
                filterednew_df=filterednew_df[filterednew_df[column_name].isin(input_json[column_name])]


    return filterednew_df

def get_pointslist(filtered_df):
    points_list=[]
    columns_G_AB=tfo_config.columns_G_AB 
    for index, excel_row in filtered_df.iterrows():
        number_of_patients=int(excel_row["PATIENTS"])
        providing_organization=excel_row["PROVIDING_ORGANIZATION"]
        study_code=excel_row["STUDY_CODE"]
        excel_row=excel_row[columns_G_AB]
        tfo_cppsimilar_logger.debug("columns G to AB ------------------")
        tfo_cppsimilar_logger.debug(excel_row.values)
        values=excel_row.values            
        sum_value=sum(value  for value in values if not (isinstance(value, str) or value!=value) )
        yaxis_value=calculate_cpp(sum_value,number_of_patients)
        points_list.append([study_code,yaxis_value,providing_organization])
    return points_list

def filter_oncologydata(input_df,input_json):
    try:
        study_code_list=[]
        yaxis_list=[]
        if input_json["STUDY_CODE"]:
                filtered_df=copy.deepcopy(input_df)
                filtered_df=filterdf_values(input_df,filtered_df, input_json,studycode_flag=True)
        else:
            filtered_df=copy.deepcopy(input_df)
            filtered_df=filterdf_values(input_df,filtered_df, input_json,studycode_flag=False)
        
        points_list=get_pointslist(filtered_df)
        filtereddata_points=filter_duplicaterows(points_list)
        for trailcode , trailvalue in filtereddata_points.items():
            study_code_list.append(trailcode)
            yaxis_list.append(trailvalue[1])
        return study_code_list,yaxis_list


    except Exception as e:
        tfo_cppsimilar_logger.debug(str(traceback.format_exc()))
        return [],[]


def oncology_filter_cppsimilartrails_data(input_filepath, input_json,planned_studycode, planned_cpp):
    '''
    Process input file to get chart data.
    '''
      
    input_df=read_inputfile(input_filepath)
    input_df.drop_duplicates(inplace=True)
    temp_dict=dict()
    study_code_list=[]
    yaxis_list=[]

    study_code_list,yaxis_list=filter_oncologydata(input_df,input_json)
    #add primary study code to list
    study_code_list.append(planned_studycode)
    yaxis_list.append(float(planned_cpp))

    temp_dict["x_axis"]=study_code_list
    temp_dict["y_axis"]=yaxis_list
    median_value=median(yaxis_list)
    tfo_cppsimilar_logger.debug("median value---------------------")
    tfo_cppsimilar_logger.debug(median_value)
    processed_df=pd.DataFrame.from_dict(temp_dict)
    return processed_df



def create_columnclustered_chartdata(input_df,chart_name):
    '''
    Extracts data from input df and prepares chart data.
    '''
    # define chart data 
    chart_data = ChartData()
    chart_data.categories = list(input_df["x_axis"])
    column_names=input_df.columns
    chart_data.add_series(tfo_config.seriesname_dict[chart_name],tuple(input_df["y_axis"]))
    return chart_data


def create_insert_cppsimilartrailchart(input_filepath, json_payload,  chart_name,planned_studycode,planned_cpp):
    '''
    Creates cpp similar trail data based on input json and geneartes column cluster chart.
    '''
    if json_payload["SlideName"]=="CPP_ONCOLOGY":
        processed_df=oncology_filter_cppsimilartrails_data(input_filepath, json_payload,planned_studycode,planned_cpp)
    else:
        processed_df=filter_cppsimilartrails_data(input_filepath, json_payload,planned_studycode,planned_cpp)
    # define chart data 
    chart_data =create_columnclustered_chartdata(processed_df,chart_name)
    #create chart
    series_number_dict=dict()
    
    return chart_data, series_number_dict

