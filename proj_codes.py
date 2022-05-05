# -*- coding: utf-8 -*-
"""
Created on Thu Apr 21 15:13:28 2022

@author: WB563112
"""

class Sectors():
    
    #Initialize the object
    def __init__(self):
        self.__data = None
        self.__dataloaded = False
        self.__last_command = None
        self.__last_output = None
        self.__last_output_exist = False
        print("Sectors object created.")
    ################################################
    
    def __bool__(self):
        """
        Description
        -------
        Boolean representation of a Sectors object.
        Returns True if data has been loaded to a Sectors object.
        """
        return self.__dataloaded
    ################################################
    
    def __str__(self):
        """
        Description
        -------
        String representation of a Sectors object.
        """
        return ("Object class: Sectors. " +
                f"Holds data: {self.__dataloaded}. " +
                f"Most recent call: {self.__last_command}.")
    ################################################

    
    def load_data(self):
        
        """
        Description
        ----------
        Loads all available data on WB project sectors.
        Data loading occurs in place and does not support variable assignment.
        User access to IEG N:\ drive is required for successful execution. 
        Data loading typically takes 1-2 minutes.
        Loading time over 3 minutes is an indication that attempt to access to IEG N:\ drive has failed.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        Creates a DataFrame object that is loaded into the private data attribute of a Sectors object.
        To access this DataFrame, use data_info or copy_data.
        """
        
        #Record command call in the last_command attribute
        self.__last_command = "load_data"
        
        #if data has already been loaded to the object
        if self.__dataloaded == True:
            print("Data already loaded to this object.")
        
        #if data not yet loaded to the object, begin data loading sequence
        else:
            
            #record start time
            import time
            start_time = time.time()                             
            
            #alert user that sequence to load data has begun
            print("Loading WB project sectors data." + "\n" +  
                  "This typically takes between 1-2 minutes. Please wait...")  
            
            #specify the name and file path of the sectors data to be imported
            import os
            import pandas as pd   
            
            #list all the files in the N drive folder, and choose the relevant file to import
            file_list = os.listdir("N:\\BASE_DATA")
            data_file = [x for x in file_list if "Project_data" in x][0]
            data_file_info = data_file.split('.')
            data_file_to_import = "N:\\BASE_DATA\\"+data_file
            
            #import projects metadata and drop unnecessary columns
            meta_data = pd.read_excel(data_file_to_import, sheet_name="metadata")  
            meta_data.drop(['Project Status Code','Lending Instrument Code'], axis=1, inplace=True)                        
            
            #import the sector data
            sector_data = pd.read_excel(data_file_to_import, sheet_name="sectors",     
                          usecols=['Project Id', 'Major Sector Code', 'Major Sector Long Name', 'Sector Code', 
                                   'Sector Long Name', 'Sector Percentage'])
            
            #merge the two datasets and compute sector percentage as a percentage
            sector_and_meta = sector_data.merge(meta_data, on="Project Id", how="outer", validate="m:1")
            sector_and_meta['Sector Percentage'] = sector_and_meta['Sector Percentage'] * 100
            
            #save the merged data to the data attribute of the object and set .dataloaded to True
            self.__data = sector_and_meta                                    
            self.__dataloaded = True                                    
            
            #delete residual files
            del(meta_data)
            del(sector_data)
            del(data_file_to_import)
            
            #record end time and calculate elapsed time
            end_time = time.time() 
            elapsed_time = end_time-start_time
            
            #Alert users that data was loaded successfully
            print("Data loading successful!" + "\n" +
                  f"Loaded data contains {self.__data.shape[0]} rows and {self.__data['Project Id'].nunique()} unique WB projects." + "\n" +
                  f"Total loading time: {round(elapsed_time, 1)} seconds." + "\n" +
                  "Data source: World Bank PowerBI Data Platform." + "\n" +
                  f"Data download date: {data_file_info[2]}.")
    ################################################        
    
    def unload_data(self): 
        """
        Description
        ----------
        Deletes the DataFrame, if any, that has been loaded into the data attribute of a Sectors object.
        Data unloading occurs in place and does not support variable assignment.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        None
        """
        
        #Record command call in the last_command attribute
        self.__last_command = "unload_data"
        
        #Delete internal data and reset all internal attributes
        self.__data = None
        self.__dataloaded = False
        self.__last_output = None
        self.__last_output_exist = False
        print("Data unloading complete.") 
    ################################################
         
    def data_info(self):
        """
        Description
        ----------
        Prints summary information on the DataFrame, if any, that has been loaded into the data attribute of a Sectors object.
        Variable assignment not supported.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        None
        """
        
        #Record command call in the last_command attribute
        self.__last_command = "data_info"
        
        #if data has been loaded to the object, return info of the df
        if self.__dataloaded:                      
            print(self.__data.info())            
        
        #Otherwise, alert user that data has not been loaded yet
        else:                                      
            print("Data not yet loaded.")                        
    ################################################
    
    def copy_data(self):
        """
        Description
        ----------
        Returns a copy of the DataFrame, if any, that has been loaded into the data attribute of a Sectors object.
        Supports variable assignment.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        DataFrame object
        """
        
        #Record command call in the last_command attribute
        self.__last_command = "copy_data"
        
        #if data has been loaded to the object, return the data
        if self.__dataloaded:                 
            return self.__data
        
        #Otherwise, alert user that data has not been loaded yet                      
        else:                                       
            print("Data not yet loaded.")
    ################################################       
    
    def get_projects(self, 
                     sector_codes,
                     min_pct=1,
                     start_FY=None, stop_FY=None,
                     product_type=None, 
                     project_status=None,
                     include_AF=True,
                     show_all=False,
                     show_meta=False):
        
        """
        Description
        ----------
        Returns projects that are mapped to the specified sector codes and that satisfy other optional criteria.
        Supports a variable assignment.
        Data must already be loaded into the Sectors object.
        
        Parameters
        ----------
        sector_codes : list
            Returned projects must be mapped to one or more sector codes in sector_codes.
            Only sub-sector level codes are accepted.
        min_pct : int, default 1
            Indicates the minimum percentage of a project that at least one sector code in sector_codes must account for before that project will be returned.
            If project's sector codes matches one or more sector codes in sector_codes but none of the matching sector codes account for up to the min_pct value, such a project will not be returned. 
        start_FY : int or None, default None
            If int, returns projects approved in or after the specified value of start_FY, and projects with missing approval FY data.
            If None, returns all eligible projects, regardless of approval FY.
        stop_FY : int or None, default None
            If int, returns projects approved in or before the specified value of stop_FY, and projects with missing approval FY data.
            If None, returns all eligible projects, regardless of approval FY.
        product_type : list or None, default None
            If list, returns projects whose product type matches any of the specified values in product_types.
            Automatically excludes projects with missing product type data.
            Acceptable values are:
                "L": for lending products,
                "A": for AAA products,
                "S": for standard products, or any combination of these. 
            If None, returns all eligible projects, regardless of product type.
        project_status : list or None, default None
            If list, returns projects whose completion status matches any of the specified values in project_status.
            Automatically excludes observations with missing project status data.
            Acceptable values are: "Active", "Canceled", "Closed", "Dropped", "Draft", "Legacy Dropped", "Legacy", and "Pipeline".
            If None, return all eligible projects regardless of project status.
        include_AF : bool, default True
            If True, eligible additional financing projects will be included in the output.
        show_all : bool, default False
            If True, returns all sector codes that matching projects are mapped to, rather than those in sector_codes only.
        show_meta : bool, default False
            If True, returns additional project-level meta data.
        
        Returns
        ----------
        DataFrame object
        """ 
        
        #Record command call in the last_command attribute
        self.__last_command = "get_projects"
        
        #if data is not loaded to the object, alert user
        if self.__dataloaded==False:
            self.__last_output = None
            self.__last_output_exist = False
            print ("Data not yet loaded.")
    
        #Otherwise, begin data extraction sequence
        else:
            
            #create copy of internal data
            temp_data = self.__data   
            
            #USER INPUT VALIDATION
            import numpy as np
            
            #Check if any auxiliary argument is provided
            aux_args = (start_FY==None) & (stop_FY==None) & (product_type==None) & (project_status==None) & (include_AF==True)
            
            #---------------------------------------------#
            #If sector_codes is not a list, return error
            if type(sector_codes)!=list:
                raise TypeError("'sector_codes' must be of type 'list'.")
            #If any item in sector_codes is not a string, return error
            elif any(type(item)!=str for item in sector_codes):
                raise TypeError("All items in 'sector_codes' must be of type 'str'.")
            #If sector_codes is a list and every element in that list is a string, convert all element in sector_codes to upper case
            else:
                sector_codes = [item.upper() for item in sector_codes]
                
            #If any element in sector_codes is a major sector code (i.e. end with 'X'), warn user
            if any(item[-1]=='X' for item in sector_codes):
                print("WARNING! One or more values in 'sector_codes' is not a sub-sector level code")
            
            #---------------------------------------------#
            #If min_pct is not an int, return error
            if type(min_pct)!=int:
                raise TypeError("'min_pct' must be of type 'int'.") 
            #If min_pct is <0 or >100, warn user
            elif min_pct<0 or min_pct>100:
                print("WARNING! 'min_pct' is outside expected range of 0 to 100.")
                
            #--------------------------------------------#
            #If start_FY is specified by user and it is not of type 'int', return error
            if (start_FY!=None) and (type(start_FY)!=int):
                raise TypeError("start_FY must be of type 'int'.")    
            
            #If stop_FY is specified by user and it is not of type 'int', return error
            if (stop_FY!=None) and (type(stop_FY)!=int):
                raise TypeError("stop_FY must be of type 'int'.")    
            
            #if start_FY and stop_FY are unspecified by user, set them to first and last year available in the data respectively
            start_FY = int(temp_data['Project Approval FY'].min()) if start_FY==None else int(start_FY)
            stop_FY = int(temp_data['Project Approval FY'].max()) if stop_FY==None else int(stop_FY)
            
            #If stop_FY precedes start_FY, return error
            if stop_FY < start_FY:
                raise ValueError("'start_FY' and 'stop_FY' arguments are not in chronological order.")
            
            #create a list of integers between start_FY and stop_FY 
            timeline = list(range(int(start_FY), int(stop_FY)+1))
            timeline.append(np.nan)
            
            #--------------------------------------------#            
            #If product_type is specified by user and it is not of type 'list', return error
            if (product_type!=None) and (type(product_type)!=list):
                raise TypeError("product_type must be of type 'list'.")
            #If product_type is a list and any element in product_type is not a string, return error
            elif (type(product_type)==list) and (any(type(item)!=str for item in product_type)):
                raise TypeError("All values in 'product_type' must be of type 'str'.")
            else:
                #If product_type is specified correctly as a list, convert to upper, else leave as None
                #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
                product_type = [item.upper() for item in product_type] if product_type!=None else product_type
                
            #Create a list of all available values for Product Line Type
            prod_type_options = list(temp_data['Product Line Type'].unique())
            #Remove any nan from this list, and record it in the nan_removed variable
            nan_removed = False
            if np.nan in prod_type_options:
                prod_type_options.remove(np.nan)
                nan_removed = True
            #Then convert prod_type_options to upper case to match product_type
            #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
            prod_type_options = [item.upper() for item in prod_type_options]
    
            #If product_type is specified by user and it is not among acceptable options, exit
            if (product_type!=None) and set(product_type).isdisjoint(set(prod_type_options)):
                raise ValueError("Unrecognized product_type input. Acceptable values are:" + "\n" +
                        "'L', for lending products," + "\n" +
                        "'A', for AAA products, and" + "\n" + 
                        "'S', for standard products.") 
            
            #If product_type is not specified by user, set prod_type to all acceptable options
            if product_type==None:
                #If nan values were removed from prod_type_options, put them back
                if nan_removed:
                    prod_type_options.append(np.nan)
                prod_type = prod_type_options
            #Finally, if product_type is specified correctly by user, set prod_type to the specified value of product_type
            else:
                prod_type = product_type
            
            #--------------------------------------------#
            #If project_status is specified by user and it is not of type 'list', return error
            if (project_status!=None) and (type(project_status)!=list):
                raise TypeError("'project_status' must be of type 'list'.")
            #If project_status is specified as a list and any item in project_status is not a string, return error
            elif (type(project_status)==list) and (any(type(item)!=str for item in project_status)):
                raise ValueError("All values in 'project_status' must be of type 'str'.")
            else:
                #If project_status is specified correctly as a list, convert to title, else leave as None
                #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
                project_status = [item.title() for item in project_status] if project_status!=None else project_status
            
            #Get the list of possible options for project status
            proj_stat_options = list(temp_data['Project Status Name'].unique())
            #Remove any nan from this list, and record it in the nan_removed variables
            nan_removed = False
            if np.nan in proj_stat_options:
                proj_stat_options.remove(np.nan)
                nan_removed=True
            #Then convert proj_stat_options to title case to match project_status
            proj_stat_options = [item.title() for item in proj_stat_options]
            
            #If project_status is specified by user and it is not among acceptable options, return error
            if (project_status!=None) and set(project_status).isdisjoint(set(proj_stat_options)):
                raise ValueError("Unrecognized project_status input. Acceptable values are:" + "\n" +
                        '"Active", "Canceled", "Closed", "Dropped", "Draft", "Legacy Dropped", "Legacy", and "Pipeline".')
            
            #If project_status is not specified by user, set proj_status to all acceptable options
            if project_status==None:
                #If nan was removed, put them back
                if nan_removed:
                    proj_stat_options.append(np.nan)
                proj_status = proj_stat_options
            #Finally, if project_status is specified correctly by user, set proj_status to the specified value in project_status
            else:
                proj_status = project_status
                
            #--------------------------------------------#
            #If include_AF is not of type 'bool', return error
            if type(include_AF)!=bool:
               raise TypeError("'include_AF' must be of type 'bool'.")
            #If include_AF is True, set add_fin_choice to all available values of the Additional Financing Flag

            add_fin_choice=list(temp_data['Additional Financing Flag'].unique())
            #If include_AF is False, exclude add_fin_choice value from the Additional Financing Flag
            if include_AF==False:
                #NOTE FOR DEVELOPER: CONFIRM THAT 'Y' IS THE VALUE FOR ADDITIONAL FINANCING FLAG
                add_fin_choice.remove('Y')
            
            #If show_all is not of type 'bool', return error
            if type(show_all)!=bool:
                raise TypeError("'show_all' must be of type 'bool'.")
                
            #If show_meta is not of type 'bool', return error
            if type(show_meta)!=bool:
                raise TypeError("'show_meta' must be of type 'bool'.")
            
            #IDENTIFY PROJECTS MATCHING THE SECTOR_CODES AND MIN_PCT ARGUMENTS SPECIFIED
            #--------------------------------------------#
            #Step 1: get the rows with sector codes that match any of those in sector_codes
            matching_sector_filter = temp_data['Sector Code'].isin(sector_codes)

            #filter temp data to extract these matching sector codes only
            matching_sector_df = temp_data.loc[matching_sector_filter,:].copy()
            
            #Step 2: Extract PIDs where at least one sector code match has a sector percentage > the min_pct value
            #Pivot the data so that each pid occupy a single row only
            matching_sector_pivoted_df = matching_sector_df.pivot(index='Project Id', 
                                                                columns='Sector Code', 
                                                                values='Sector Percentage')
            
            #compute a min_pct_exceeded column that returns True if at least on sector code has percentage >= min_pct
            matching_sector_pivoted_df['min_pct_exceeded'] = (matching_sector_pivoted_df>=min_pct).any(axis=1,skipna=True)
            
            #Reset index
            matching_sector_pivoted_df.reset_index(inplace=True)
            
            #Extract PIDs where min_pct_exceeded column is True
            pids_with_sector_above_min_pct = list(matching_sector_pivoted_df.loc[matching_sector_pivoted_df['min_pct_exceeded'],'Project Id'])
            
            #--------------------------------------------#
            #Specify the output rows depending on the values of the auxiliary arguments
            if aux_args:
                output_rows = temp_data['Project Id'].isin(pids_with_sector_above_min_pct)
            else:
                output_rows = ((temp_data['Project Id'].isin(pids_with_sector_above_min_pct)) &            
                               (temp_data['Product Line Type'].isin(prod_type)) &
                               (temp_data['Project Status Name'].isin(proj_status)) &
                               (temp_data['Project Approval FY'].isin(timeline)) & 
                               (temp_data['Additional Financing Flag'].isin(add_fin_choice)))
            
            #--------------------------------------------#
            #Specify the output columns depending on the value of show_meta
            all_cols = list(temp_data.columns)
            sel_cols = ['Project Id', 'Sector Code', 'Sector Long Name', 'Sector Percentage', 'Additional Financing Flag',
                        'Project Approval FY', 'Project Status Name', 'Product Line Type', 'Lead GP/Global Themes']
            
            #Output columns is all columns if show_meta is true, else show sel_cols
            output_cols = all_cols if show_meta else sel_cols
            
            #--------------------------------------------#
            #Filter the data based on output_rows and output_cols
            output_df = temp_data.loc[output_rows,output_cols].copy()
            
            #If user only wants to see the specified sector codes rather than all sector codes the matching projects are mapped to
            if show_all==False:
                #filter the data to show only the specified sector codes
                output_df = output_df.loc[(output_df['Sector Code'].isin(sector_codes) & 
                                           (output_df['Sector Percentage']>=min_pct)),:].copy()
            #reset the index
            output_df.reset_index(drop=True,inplace=True)
            
            #--------------------------------------------#
            #Check if output df is empty and alert user accordingly
            if output_df.empty:
                print("No projects meet the specified criteria.")
            else: 
                no_of_unique = output_df['Project Id'].nunique()
                print(f"{no_of_unique} unique projects found.") 
                self.__last_output = output_df
                self.__last_output_exist = True
                return output_df
    ################################################        
    
    def get_sectors(self, pid_list, show_meta=False):
                
        """
        Description
        ----------
        Returns sector codes that the specified list of projects are mapped to.
        Supports variable assignment.
        Data must already be loaded into the Sectors object.
        
        Parameters
        ----------
        pid_list : list
            The ID numbers of projects for which sector codes are to be returned. 
        show_meta : bool, default False
            If True, returns additional project-level meta data.
        
        Returns
        ----------
        DataFrame object
        """
        #Record command call in the last_command attribute
        self.__last_command = "get_sectors"
        
        #if data is not loaded to the object, alert user
        if self.__dataloaded==False:
            self.__last_output = None
            self.__last_output_exist = False
            print ("Data not yet loaded.")

        #Otherwise, begin data extraction sequence
        else:
            
            #create copy of internal data
            temp_data = self.__data  

            #USER INPUT VALIDATION
            #If pids is not a list, return error
            if type(pid_list)!=list:
                raise TypeError("'pid_list' must be of type 'list'.")
            #If each pid in pids is not a string, return error
            elif any(type(pid_list)!=str for pid in pid_list):
                raise TypeError("Every item in 'pid_list' must be of type 'str'.")
                
            #Convert pids to upper
            #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
            pids = [item.upper() for item in pid_list]
      
            #--------------------------------------------#
            #Identify the rows in temp_data that matches the PIDs
            output_rows = temp_data['Project Id'].isin(pids)
            
            #--------------------------------------------#
            #Specify the output columns depending on the value of show_meta
            all_cols = list(temp_data.columns)
            sel_cols = ['Project Id', 'Major Sector Code', 'Major Sector Long Name', 'Sector Code', 'Sector Long Name', 'Sector Percentage']
            
            #Show all columns if show_meta is true, else show sel_cols
            output_cols = all_cols if show_meta else sel_cols
        
            #--------------------------------------------#
            #Filter the data based on output_rows and output_cols
            output_df = temp_data.loc[output_rows,output_cols].copy()
            output_df.reset_index(drop=True,inplace=True)
            
            #--------------------------------------------#
            #Check if output df is empty and alert user accordingly
            if output_df.empty:
                print("No sector codes found.")
            else:
                no_of_unique = output_df['Project Id'].nunique()
                print(f"Sector codes found for {no_of_unique} out of {len(pids)} requested projects.")
                self.__last_output = output_df
                self.__last_output_exist = True
                return output_df
    ################################################      
            
    def save_last(self, save_name=None):
 
        """
        Description
        ----------
        Exports the output from the most recent call of a get_projects or get_sectors command.
        Output is saved in .xlsx format.
        
        Parameters
        ----------
        save_name : str or None, default None
            Name exported data should be saved with.
            If None, exported data is saved with name "Sector_extract".
            Overwrites any pre-existing file of the same name in the save folder.
        
        Returns
        ----------
        None
        """
        #Record command in the last_command attribute
        self.__last_command = "save_last"
        
        #If there is no last output, alert user
        if self.__last_output_exist == False:
            print("No output to save.")
        #If there is last output
        else:
            #Check if save_name is specified
            if save_name==None:
                save_name = "Sector_extract.xlsx"
            else:
                #If specified save_name is not a str, return error
                if type(save_name)!=str:
                    raise TypeError("'save_name' must be of type 'str'.")
                #Add fill extension to the save_name
                save_name = save_name + ".xlsx"
            #Import pandas
            import pandas as pd
            #Extract the output df
            output_df = self.__last_output 
            #Save it to save_name
            output_df.to_excel(save_name,index=False)
            print("Output saved.")
    ################################################       
    
    def plot_gp(self, save_it=False, save_name=None):
        """
        Description
        ----------
        Plots the distribution of project by Global Practice, using output from the most recent call of a get_projects command.
        Must be called immediately after the corresponding get_projects or plot_fy commands. 
        
        Parameters
        ----------
        save_it : bool, default None
            If True, plot created is saved in .png format. 
        save_name : str or None, default None
            The name plot created should be saved with.
            If None, plot is saved with name "Sectors_by_GP".
            Overwrites any pre-existing file of the same name in the save folder.
        
        Returns
        ----------
        None
        """
        
        import sys
        
        #If save_it is not of type boolean, exit
        if type(save_it)!= bool:
           raise TypeError("save_it must be a boolean.") 
        
        #If the previous command is not the get_projects command, return error
        if ((self.__last_command!= "get_projects") and 
            (self.__last_command!="plot_fy") and 
            (self.__last_command!="plot_gp") and 
            (self.__last_command!="save_last")):
           sys.exit("Command not preceded by 'get_projects'.") 
        
        #Record command in the last_command attribute
        self.__last_command = "plot_gp"
        
        #If the previous get_projects command did not return a valid output, alert user
        if self.__last_output_exist == False:
            print("No output to plot.")
            
        else:
            
            #If save_name is none, set it to a default value
            if save_name==None:
                save_name = "Sectors_by_GP"
            else:
                #If save_name is specified and it is not a str, return error
                if type(save_name)!=str:
                    raise TypeError("'save_name' must be of type 'str'.")
                    
            #Import plt
            from matplotlib import pyplot as plt
            
            #Extract the output_df from the last get_projects command
            output_df = self.__last_output
            
            #Drop duplicates on Project Id
            unique_df = output_df.drop_duplicates(subset="Project Id", keep="first")
            #Count number of Project Ids missing GP data
            missing_FY_count = unique_df["Lead GP/Global Themes"].isna().sum()
            #If missing_FY_count > 0, notify user of of the impending exclusion
            if missing_FY_count>0:
                print(f"WARNING! {missing_FY_count} project(s) with missing GP data got excluded from the plot.")
            
            #Group the output_df by GP and count the number of unique project ID
            plot_series = output_df.groupby("Lead GP/Global Themes")["Project Id"].nunique()
            
            #Reset the index to transform plot series to dataframe
            plot_df = plot_series.reset_index()
            
            #Sort data
            plot_df.sort_values("Project Id", ascending=True, inplace=True)
            
            #Create plot canvass
            fig, ax = plt.subplots()
            
            #Create horizontal bar chart, using a container to save the chart
            fig_con = ax.barh(plot_df["Lead GP/Global Themes"],plot_df["Project Id"])
            
            #Add data label
            ax.bar_label(fig_con)
            
            # make the x ticks integers, not floats
            x_label_int = []
            locs, labels = plt.xticks()
            for each in locs:
                x_label_int.append(int(each))
            plt.xticks(x_label_int)
            
            #Add title to X-Axis
            ax.set_xlabel("Number of projects")
            
            #Add title to Y-Axis
            ax.set_ylabel("Lead GP/Global Themes")
            
            #Add chart title
            ax.set_title("Project count, by Global Practice")
            
            if save_it:
                print("Plot saved.")
            
            fig.savefig(save_name, bbox_inches='tight') if save_it else plt.show()
    ################################################            
    
    def plot_fy(self, save_it=False, save_name=None):
        """
        Description
        ----------
        Plots the distribution of project by Approval Fiscal Year, using the output from the most recent call of a get_projects command.
        Must be called immediately after the corresponding get_projects or plot_gp commands. 
        
        Parameters
        ----------
        save_it : bool, default None
            If True, plot created is saved in .png format. 
        save_name : str or None, default None
            The name plot created should be saved with.
            If None, plot is saved with name "Sectors_by_FY".
            Overwrites any pre-existing file of the same name in the save folder.
        
        Returns
        ----------
        None
        """
        
        import sys
        
        #If save_it is not of type boolean, exit
        if type(save_it)!= bool:
           raise TypeError("'save_it' must be a boolean.") 
        
        #If the previous command is not the get_projects command, exit
        if ((self.__last_command!="get_projects") and 
            (self.__last_command!="plot_gp") and 
            (self.__last_command!="plot_fy") and 
            (self.__last_command!="save_last")):
           sys.exit("Command not preceded by 'get_projects'.") 
        
        #Record command in the last_command attribute
        self.__last_command = "plot_fy"
        
        #If the previous get_projects command did not return a valid output, alert user
        if self.__last_output_exist == False:
            print("No output to plot.")
            
        else:
            
            #If save_name is none, set it to a default value
            if save_name==None:
                save_name = "Sectors_by_FY"
            else:
                #If save_name is specified and it is not a str, return error
                if type(save_name)!=str:
                    raise TypeError("'save_name' must be of type 'str'.")
                    
            #Import plt
            from matplotlib import pyplot as plt
            
            #Extract the output_df from the last get_projects command
            output_df = self.__last_output
            
            #Drop duplicates on Project Id
            unique_df = output_df.drop_duplicates(subset="Project Id", keep="first")
            #Count number of Project Ids missing Approval FY data
            missing_FY_count = unique_df["Project Approval FY"].isna().sum()
            #If missing_FY_count > 0, notify user of of the impending exclusion
            if missing_FY_count>0:
                print(f"WARNING! {missing_FY_count} project(s) with missing Approval FY data got excluded from the plot.")
            
            #Group the output_df by GP and count the number of unique project ID
            plot_series = output_df.groupby("Project Approval FY")["Project Id"].nunique()
            
            #Reset the index to transform plot series to dataframe
            plot_df = plot_series.reset_index()
            
            #Sort data
            plot_df.sort_values("Project Approval FY", ascending=True, inplace=True)
            
            #Create plot canvass
            fig, ax = plt.subplots()
            
            #Create horizontal bar chart, using a container to save the chart
            fig_con = ax.bar(plot_df["Project Approval FY"],plot_df["Project Id"])
            
            #Add data label
            ax.bar_label(fig_con)
            
            #Add title to X-Axis
            ax.set_xlabel("Project Approval FY")
            
            #Add title to Y-Axis
            ax.set_ylabel("Number of projects")
            
            #Add chart title
            ax.set_title("Project count, by Project Approval FY")
            
            if save_it:
                print("Plot saved.")
            
            fig.savefig(save_name, bbox_inches='tight') if save_it else plt.show()
    ################################################  
    ################################################
    ################################################

class Themes():
    
    #Initialize the object
    def __init__(self):
        self.__data = None
        self.__dataloaded = False
        self.__last_command = None
        self.__last_output = None
        self.__last_output_exist = False
        print("Themes object created.")
    ################################################
    
    def __bool__(self):
        """
        Description
        -------
        Boolean representation of a Themes object.
        Returns True if data has been loaded to a Themes object.
        """
        return self.__dataloaded
    ################################################
    
    def __str__(self):
        """
        Description
        -------
        String representation of a Themes object.
        """
        return ("Object class: Themes. " +
                f"Holds data: {self.__dataloaded}. " +
                f"Most recent call: {self.__last_command}.")
    ################################################
    
    def load_data(self):
        
        """
        Description
        ----------
        Loads all available data on WB project themes.
        Data loading occurs in place and does not support variable assignment.
        User access to IEG N:\ drive is required for successful execution. 
        Data loading typically takes 2-4 minutes.
        Loading time over 4 minutes is an indication that attempt to access to IEG N:\ drive has failed.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        Creates a DataFrame object that is loaded into the private data attribute of a Themes object.
        To access this DataFrame, use data_info or copy_data.
        """
        
        #Record function call in the last_command attribute
        self.__last_command = "load_data"
        
        #if data has already been loaded to the object
        if self.__dataloaded == True:
            print("Data already loaded to this object.")
        
        #if data not yet loaded to the object
        else:   
            
            #record start time
            import time
            start_time = time.time()                             
            
            #let user know that sequence to load data has begun
            print("Loading WB project Themes data." + "\n" +  
                  "This typically takes between 2-5 minutes. Please wait...")  
            
            #specify the name and file path of the Themes data to be imported
            import os
            import pandas as pd   
            
            #list all the files in the N drive folder, and choose the relevant file to import
            file_list = os.listdir("N:\\BASE_DATA")
            data_file = [x for x in file_list if "Project_data" in x][0]
            data_file_info = data_file.split('.')
            data_file_to_import = "N:\\BASE_DATA\\"+data_file
            
            #import projects metadata
            meta_data = pd.read_excel(data_file_to_import, sheet_name="metadata")
            meta_data.drop(['Project Status Code','Lending Instrument Code'], axis=1, inplace=True)                         
            
            #import the Themes data
            theme_data = pd.read_excel(data_file_to_import, sheet_name="themes",     
                          usecols=['Project Id', 'Theme Code', 'Theme Level', 'Theme Name', 
                                   'Theme Percentage', 'Theme Lending Commitment Amount', 'Theme Portfolio Net Commitment Amount'])
            
            #merge the two datasets
            theme_and_meta = theme_data.merge(meta_data, on="Project Id", how="outer", validate="m:1")
            theme_and_meta['Theme Percentage'] = theme_and_meta['Theme Percentage'] * 100
            
            #save the merged data to the data attribute of the object and set .dataloaded to True
            self.__data = theme_and_meta                                    
            self.__dataloaded = True                                    
            
            #delete residual files
            del(meta_data)
            del(theme_data)
            del(data_file_to_import)
            
            #record end time and calculate elapsed time
            end_time = time.time() 
            elapsed_time = end_time-start_time
            
            #Alert users that data was loaded successfully
            print("Data loading successful!" + "\n" +
                  f"Loaded data contains {self.__data.shape[0]} rows and {self.__data['Project Id'].nunique()} unique WB projects." + "\n" +
                  f"Total loading time: {round(elapsed_time, 1)} seconds." + "\n" +
                  "Data source: World Bank PowerBI Data Platform." + "\n" +
                  f"Data download date: {data_file_info[2]}.")
    ################################################
    
    def unload_data(self):
        
        """
        Description
        ----------
        Deletes the DataFrame, if any, that has been loaded into the data attribute of a Themes object.
        Data unloading occurs in place and does not support variable assignment.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        None
        """
        
        #Record function call in the last_command attribute
        self.__last_command = "unload_data"
        
        #Delete internal data and reset all internal attributes
        self.__data = None
        self.__dataloaded = False
        self.__last_output = None
        self.__last_output_exist = False
        print("Data unloading complete.") 
    ################################################
         
    def data_info(self):
        """
        Description
        ----------
        Prints summary information on the DataFrame, if any, that has been loaded into the data attribute of a Themes object.
        Variable assignment not supported.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        None
        """
        
        #Record function call in the last_command attribute
        self.__last_command = "data_info"
        
        #if data has been loaded to the object, return info of the df
        if self.__dataloaded:                      
            print(self.__data.info())            
        
        #Otherwise, alert user that data has not been loaded yet
        else:                                      
            print("Data not yet loaded.")                        
    ################################################
    
    def copy_data(self):
        """
        Description
        ----------
        Returns a copy of the DataFrame, if any, that has been loaded into the data attribute of a WBTheme object.
        Supports variable assignment.
        
        Parameters
        ----------
        None
        
        Returns
        ----------
        DataFrame object
        """
        
        #Record function call in the last_command attribute
        self.__last_command = "copy_data"
        
        #if data has been loaded to the object, return the data
        if self.__dataloaded:                 
            return self.__data
        
        #Otherwise, alert user that data has not been loaded yet                      
        else:                                       
            print("Data not yet loaded.")
    ################################################    
        
    def get_projects(self, 
                     theme_codes,
                     min_pct=1,
                     start_FY=None, stop_FY=None,
                     product_type=None, 
                     project_status=None,
                     include_AF=True,
                     show_all=False,
                     show_meta=False):
        
        """
        Description
        ----------
        Returns projects that are mapped to the specified theme codes and that satisfy other optional criteria.
        Supports a variable assignment.
        Data must already be loaded into the Themes object.
        
        Parameters
        ----------
        theme_codes : list
        Returned projects must be mapped to one or more theme codes in theme_codes.
        
        min_pct : int, default 1
            Indicates the minimum percentage of a project that at least one theme code in theme_codes must account for before that project will be returned.
            If project's theme codes matches one or more theme codes in theme_codes but none of the matching theme codes account for up to the min_pct value, such a project will not be returned. 
        start_FY : int or None, default None
            If int, returns projects approved in or after the specified value of start_FY, and projects with missing Approval FY data.
            If None, returns all eligible projects, regardless of approval FY.
        stop_FY : int or None, default None
            If int, returns projects approved in or before the specified value of stop_FY, and projects with missing Approval FY data.
            If None, returns all eligible projects, regardless of approval FY.    
        product_type : list or None, default None
            If list, returns projects whose product type matches any of the specified values in product_types.
            Automatically excludes projects with missing product type data.
            Acceptable values are:
                "L": for lending products,
                "A": for AAA products,
                "S": for standard products, or any combination of these. 
            If None, returns all eligible projects regardless of product type.
        project_status : list or None, default None
            If list, returns projects with project status matches any of the specified values in project_status.
            Automatically excludes observations with missing project status data.
            Acceptable values are as follows: "Active", "Canceled", "Closed", "Dropped", "Draft", "Legacy Dropped", "Legacy", and "Pipeline".
            If None, return all eligible projects regardless of project status.
        show_all : bool, default False
            If True, returns all theme codes that matching projects are mapped to, rather than those that match theme_codes only.
        show_meta : bool, default False
            If True, returns additional project-level meta data.
        
        Returns
        ----------
        DataFrame object
        """ 
        
        #Record function call in the last_command attribute
        self.__last_command = "get_projects"
        
        #if data is not loaded to the object, alert user
        if self.__dataloaded==False:
            self.__last_output = None
            self.__last_output_exist = False
            print ("Data not yet loaded.")
    
        #Otherwise, begin data extraction sequence
        else:
            
            
            #create copy of internal data
            temp_data = self.__data   
            
            #USER INPUT VALIDATION
            import numpy as np
            
            #Check if any auxiliary argument is provided
            aux_args = (start_FY==None) & (stop_FY==None) & (product_type==None) & (project_status==None) & (include_AF==True)
            
            #---------------------------------------------#
            #If theme_code is not a list, return error
            if type(theme_codes)!=list:
                raise TypeError("'theme_codes' must be of type 'list'.")
            #If any item in theme_codes is not a string, return error
            elif any(type(item)!=int for item in theme_codes):
                raise TypeError("Every item in 'theme_codes' must be of type 'int'.")
                
            #---------------------------------------------#
            #If min_pct is not an int, return error
            if type(min_pct)!=int:
                raise TypeError("'min_pct' must be of type 'int'.")
            #If min_pct is <0 or >100, warn user
            elif min_pct<0 or min_pct>100:
                print("WARNING! 'min_pct' is outside expected range of 0 to 100.")
                
            #--------------------------------------------#
            #If start_FY is specified by user and it is not of type 'int', return error
            if (start_FY!=None) and (type(start_FY)!=int):
                raise TypeError("'start_FY' must be of type 'int'.")    
            
            #If stop_FY is specified by user and it is not of type 'int', return error
            if (stop_FY!=None) and (type(stop_FY)!=int):
                raise TypeError("'stop_FY' must be of type 'int'.")    
            
            #if start_FY and stop_FY are unspecified by user, set them to first and last year available in the data respectively
            start_FY = int(temp_data['Project Approval FY'].min()) if start_FY==None else int(start_FY)
            stop_FY = int(temp_data['Project Approval FY'].max()) if stop_FY==None else int(stop_FY)
            
            #If stop_FY precedes stop_FY, return error
            if stop_FY < start_FY:
                raise ValueError("Values of 'start_FY' and 'stop_FY' are not in chronological order.")
            
            #create a list of integers between start_FY and stop_FY 
            timeline = list(range(int(start_FY), int(stop_FY)+1))
            timeline.append(np.nan)
            
            #--------------------------------------------#
            #If product_type is specified by user and it is not of type 'list', return error
            if (product_type!=None) and (type(product_type)!=list):
                raise TypeError("'product_type' must be of type 'list'.")
            #If product_type is a list and any element in product_type is not a string, return error
            elif (type(product_type)==list) and (any(type(item)!=str for item in product_type)):
                raise ValueError("Every item in 'product_type' must be of type 'str'.")
            else:
                #If product_type is specified correctly as a list, convert to upper, else leave as None
                #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
                product_type = [item.upper() for item in product_type] if product_type!=None else product_type
                
            #Create a list of all available values for Product Line Type
            prod_type_options = list(temp_data['Product Line Type'].unique())
            #Remove any nan from this list, and record it in the nan_removed variable
            nan_removed = False
            if np.nan in prod_type_options:
                prod_type_options.remove(np.nan)
                nan_removed = True
            #Then convert prod_type_options to upper case to match product_type
            #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
            prod_type_options = [item.upper() for item in prod_type_options]
    
            #If product_type is specified by user and it is not among acceptable options, return error
            if (product_type!=None) and set(product_type).isdisjoint(set(prod_type_options)):
                raise ValueError("Unrecognized product_type input. Acceptable values are:" + "\n" +
                        "'L', for lending products," + "\n" +
                        "'A', for AAA products, and" + "\n" + 
                        "'S', for standard products.") 
            
            #If product_type is not specified by user, set product_type to all acceptable options
            if product_type==None:
                #If nan values were removed from prod_type_options, put them back
                if nan_removed:
                    prod_type_options.append(np.nan)
                prod_type = prod_type_options
            
            #Finally, if product_type is specified correctly by user, set product_type to the corresponding value in prod_type_options
            else:
                prod_type = product_type
                
            #--------------------------------------------#
            #If project_status is specified by user and it is not of type 'list', return error
            if (project_status!=None) and (type(project_status)!=list):
                raise TypeError("'project_status' must be of type 'list'.")
            #If project_status is specified as a list and any item in project_status is not a string, return error
            elif (type(project_status)==list) and (any(type(item)!=str for item in project_status)):
                raise TypeError("Every item in 'project_status' must be of type 'str'.")
            else:
                #If project_status is specified correctly as a list, convert to title, else leave as None
                #NOTE FOR DEVELOPER: CHECK THAT CASING STYLE USED MATCHES THAT IN THE RAW DATA FROM POWERBI
                project_status = [item.title() for item in project_status] if project_status!=None else project_status
                
            #Get the list of possible options for project status
            proj_stat_options = list(temp_data['Project Status Name'].unique())
            #Remove any nan from this list, and record it in the nan_removed variables
            nan_removed = False
            if np.nan in proj_stat_options:
                proj_stat_options.remove(np.nan)
                nan_removed=True
            #Then convert proj_stat_options to title case to match project_status
            proj_stat_options = [item.title() for item in proj_stat_options]
            
            #If project_status is specified by user and it is not among acceptable options, return error
            if (project_status!=None) and set(project_status).isdisjoint(set(proj_stat_options)):
                raise ValueError("Unrecognized project_status input. Acceptable values are:" + "\n" +
                        '"Active", "Canceled", "Closed", "Dropped", "Draft", "Legacy Dropped", "Legacy", and "Pipeline".')
            
            #If project_status is not specified by user, set product_type to all acceptable options
            if project_status==None:
                #If nan was removed, put them back
                if nan_removed:
                    proj_stat_options.append(np.nan)
                proj_status = proj_stat_options           
            #Finally, if project_status is specified correctly by user, set proj_status to the specified value in project_status
            else:
                proj_status = project_status
                
            #--------------------------------------------#
            #If include_AF is not of type 'bool', return error
            if type(include_AF)!=bool:
                raise TypeError("'include_AF' must be of type 'bool'.")
            #If include_AF is True, set add_fin_choice to all available values of the Additional Financing Flag
            elif include_AF==True:
                add_fin_choice=list(temp_data['Additional Financing Flag'].unique())
            #If include_AF is False, exclude add_fin_choice value from the Additional Financing Flag
            else:
                add_fin_choice=list(temp_data['Additional Financing Flag'].unique())
                #NOTE FOR DEVELOPER: CONFIRM THAT 'Y' IS THE VALUE FOR ADDITIONAL FINANCING FLAG
                add_fin_choice.remove('Y')
            
            #--------------------------------------------#
            #If show_all is not of type 'bool', return error
            if type(show_all)!=bool:
                raise TypeError("show_all must be of type 'bool'.")
                
            #If show_meta is not of type 'bool', return error
            if type(show_meta)!=bool:
                raise TypeError("show_meta must be of type 'bool'.")
            
            #IDENTIFY PROJECTS MATCHING THE TARGET_THEMES AND Min_PCT ARGUMENTS SPECIFIED
            #--------------------------------------------#
            #Step 1: get the rows with sector codes that match any of those in theme_codes
            matching_theme_filter = temp_data['Theme Code'].isin(theme_codes)

            #filter temp data to extract these matching sector codes only
            matching_theme_df = temp_data.loc[matching_theme_filter,:].copy()
            
            #Step 2: Extract PIDs where at least one sector code match has a theme percentage >= the min_pct value
            #Pivot the data so that each pid occupy a single row only
            matching_theme_pivoted_df = matching_theme_df.pivot(index='Project Id', 
                                                                columns='Theme Code', 
                                                                values='Theme Percentage')
            
            #compute a min_pct_exceeded column that returns True if at least on theme code has percentage >= min_pct
            matching_theme_pivoted_df['min_pct_exceeded'] = (matching_theme_pivoted_df>=min_pct).any(axis=1,skipna=True)
            
            #Reset index
            matching_theme_pivoted_df.reset_index(inplace=True)
            
            #Extract PIDs where min_pct_exceeded column is True
            pids_with_theme_above_min_pct = list(matching_theme_pivoted_df.loc[matching_theme_pivoted_df['min_pct_exceeded'],'Project Id'])
            
            #--------------------------------------------#
            #Specify the output rows depending on the values of the auxiliary arguments
            if aux_args:
                output_rows = temp_data['Project Id'].isin(pids_with_theme_above_min_pct)
            else:
                output_rows = ((temp_data['Project Id'].isin(pids_with_theme_above_min_pct)) &            
                               (temp_data['Product Line Type'].isin(prod_type)) &
                               (temp_data['Project Status Name'].isin(proj_status)) &
                               (temp_data['Project Approval FY'].isin(timeline)) & 
                               (temp_data['Additional Financing Flag'].isin(add_fin_choice)))
            
            #--------------------------------------------#
            #Specify the output columns depending on the value of show_meta
            all_cols = list(temp_data.columns)
            sel_cols = ['Project Id', 'Theme Code', 'Theme Name', 'Theme Percentage', 
                        'Project Approval FY', 'Project Status Name', 'Product Line Type', 
                        'Additional Financing Flag', 'Lead GP/Global Themes']
            
            #Output columns is all columns if show_meta is true, else show sel_cols
            output_cols = all_cols if show_meta else sel_cols
            
            #--------------------------------------------#
            #Filter the data based on output_rows and output_cols
            output_df = temp_data.loc[output_rows,output_cols]
            
            #If user only wants to see the specified theme codes rather than all theme codes the matching projects are mapped to
            if show_all==False:
                #filter the data to show only the selected theme codes at cut-off
                output_df = output_df.loc[output_df['Theme Code'].isin(theme_codes) & 
                                           (output_df['Theme Percentage']>=min_pct), :].copy()
                #reset the index
                output_df.reset_index(drop=True,inplace=True)
            
            #--------------------------------------------#
            #Check if output df is empty and alert user accordingly
            if output_df.empty:
                print("No projects meet the specified criteria.")
            else: 
                no_of_unique = output_df['Project Id'].nunique()
                print(f"{no_of_unique} unique projects found.") 
                self.__last_output = output_df
                self.__last_output_exist = True
                return output_df
    ################################################        
                  
    def get_themes(self, pid_list, theme_level=None, show_meta=False):
                
        """
        Description
        ----------
        Returns theme codes that the specified projects are mapped to.
        Supports variable assignment.
        Data must already be loaded into the Themes object.
        
        Parameters
        ----------
        pid_list : list
            The ID numbers of projects for which theme codes are to be returned. 
        show_meta : bool, default False
            If True, returns additional project-level meta data.
        theme_level : list or None, default None
            If list, returns only the theme codes matching the specified level(s).
            Acceptable values are:
                1: for the highest theme level,
                2: for middle theme level,
                3: for lowest theme level, or any combination of these. 
            If None, returns all eligible themes, regardless of theme level. 
        
        Returns
        ----------
        DataFrame object
        """
        #Record function call in the last_command attribute
        self.__last_command = "get_themes"
        
        #if data is not loaded to the object, alert user
        if self.__dataloaded==False:
            self.__last_output = None
            self.__last_output_exist = False
            print ("Data not yet loaded.")

        #Otherwise, begin data extraction sequence
        else:
            
            #create copy of internal data
            temp_data = self.__data  

            #USER INPUT VALIDATION
            #If pid_list is not a list, return error
            if type(pid_list)!=list:
                raise TypeError("'pid_list' must be of type 'list'.")
                
            #--------------------------------------------#
            #Create a list of all available values for Product Line Type
            theme_level_options = [1,2,3]
             
            #If theme_level is specified by user and it is not of type 'list', return error
            if (theme_level!=None) and (type(theme_level)!=list):
                 raise TypeError("'theme_level' must be of type list'.")
            #If theme_level is specified by user and it is not among acceptable options, return error
            elif (theme_level!=None) and set(theme_level).isdisjoint(set(theme_level_options)):
                 raise ValueError("Unrecognized theme_level input. Acceptable values are: 1, 2, and 3.")  
            #If theme_level is not specified by user, set product_type to all acceptable options
            elif theme_level==None:
                 level = theme_level_options 
            #Finally, if theme_level is specified correctly by user, set product_type to the corresponding value in prod_type_options
            else:
                 level = theme_level
                    
            #--------------------------------------------#
            #Identify the rows in temp_data that matches the PIDs
            output_rows = ((temp_data['Project Id'].isin(pid_list)) & 
                           temp_data['Theme Level'].isin(level))
            
            #--------------------------------------------#
            #Specify the output columns depending on the value of show_meta
            all_cols = list(temp_data.columns)
            sel_cols = ['Project Id', 'Theme Code', 'Theme Name', 'Theme Percentage']
            
            #Show all columns if show_meta is true, else show sel_cols
            output_cols = all_cols if show_meta else sel_cols
        
            #--------------------------------------------#
            #Filter the data based on output_rows and output_cols
            output_df = temp_data.loc[output_rows,output_cols].copy()
            output_df.reset_index(drop=True,inplace=True)
            
            #--------------------------------------------#
            #Check if output df is empty and alert user accordingly
            if output_df.empty:
                print("No theme codes found.")
            else:
                no_of_unique = output_df['Project Id'].nunique()
                print(f"Theme codes found for {no_of_unique} out of {len(pid_list)} requested projects.")
                self.__last_output = output_df
                self.__last_output_exist = True
                return output_df 
    ################################################
        
    def save_last(self, save_name=None):
        
        """
        Description
        ----------
        Exports the output from the most recent call of a get_projects or get_themes command.
        Output is saved in .xlsx format.
        
        Parameters
        ----------
        save_name : str or None, default None
            Name exported data should be saved with.
            If None, exported data is saved with name "Themes_extract".  
            Overwrites any pre-existing file of the same name in the save folder
        
        Returns
        ----------
        None
        """
        #Record function call in the last_command attribute
        self.__last_command = "save_last"
        
        #If there is no last output, alert user
        if self.__last_output_exist == False:
            print ("No output to save.")
        #If there is last output
        else:
            #Check if save_name is specified
            if save_name==None:
                save_name = "Themes_extract.xlsx"
            else:
                #If specified save_name is not a str, return error
                if type(save_name)!=str:
                    raise TypeError("'save_name' must be of type 'str'.")
                #Add fill extension to the save_name
                save_name = save_name + ".xlsx"
            #Import pandas
            import pandas as pd
            #Extract the output df
            output_df = self.__last_output 
            #Save it to save_name
            output_df.to_excel(save_name,index=False)
            print("Output saved.")
    ################################################           
            
    def plot_gp(self, save_it=False, save_name=None):
        """
        Description
        ----------
        Plots the distribution of project by Global Practice, using output from the most recent call of a get_projects command.
        Must be called immediately after the corresponding get_projects or plot_fy commands.
        
        Parameters
        ----------
        save_it : bool, default None
            If True, plot created is saved in .png format. 
        save_name : str or None, default None
            The name plot created should be saved with.
            If None, plot is saved with name "Sectors_by_GP".
            Overwrites any pre-existing file of the same name in the save folder.
        
        Returns
        ----------
        None
        """
        
        import sys
        
        #If save_it is not of type boolean, exit
        if type(save_it)!= bool:
           sys.exit("save_it must be a boolean.") 
        
        #If the previous command is not the get_projects command, exit
        if ((self.__last_command!="get_projects") and 
            (self.__last_command!="plot_fy") and 
            (self.__last_command!="plot_gp") and 
            (self.__last_command!="save_last")):
           sys.exit("Command not preceded by 'get_projects'.") 
        
        #Record function call in the last_command attribute
        self.__last_command = "plot_gp"
        
        #If the previous get_projects command did not return a valid output, alert user
        if self.__last_output_exist == False:
            print("No output to plot.")
            
        else:
            
            #If save_name is none, set it to a default value
            if save_name==None:
                save_name = "Themes_by_GP"
            else:
                #If save_name is specified and it is not a str, return error
                if type(save_name)!=str:
                    raise TypeError("'save_name' must be of type 'str'.")
                    
            #Import plt
            from matplotlib import pyplot as plt
            
            #Extract the output_df from the last get_projects command
            output_df = self.__last_output
            
            #Drop duplicates on Project Id
            unique_df = output_df.drop_duplicates(subset="Project Id", keep="first")
            #Count number of Project Ids missing GP data
            missing_FY_count = unique_df["Lead GP/Global Themes"].isna().sum()
            #If missing_FY_count > 0, notify user of of the impending exclusion
            if missing_FY_count>0:
                print(f"WARNING! {missing_FY_count} project(s) with missing GP data got excluded from the plot.")
            
            #Group the output_df by GP and count the number of unique project ID
            plot_series = output_df.groupby("Lead GP/Global Themes")["Project Id"].nunique()
            
            #Reset the index to transform plot series to dataframe
            plot_df = plot_series.reset_index()
            
            #Sort data
            plot_df.sort_values("Project Id", ascending=True, inplace=True)
            
            #Create plot canvass
            fig, ax = plt.subplots()
            
            #Create horizontal bar chart, using a container to save the chart
            fig_con = ax.barh(plot_df["Lead GP/Global Themes"],plot_df["Project Id"])
            
            #Add data label
            ax.bar_label(fig_con)
            
            # make the x ticks integers, not floats
            x_label_int = []
            locs, labels = plt.xticks()
            for each in locs:
                x_label_int.append(int(each))
            plt.xticks(x_label_int)
            
            #Add title to X-Axis
            ax.set_xlabel("Number of projects")
            
            #Add title to Y-Axis
            ax.set_ylabel("Lead GP/Global Themes")
            
            #Add chart title
            ax.set_title("Project count, by Global Practice")
            
            if save_it:
                print("Plot saved.")
            
            fig.savefig(save_name, bbox_inches='tight') if save_it else plt.show()
            
    
    def plot_fy(self, save_it=False, save_name=None):
        """
        Description
        ----------
        Plots the distribution of project by Approval Fiscal Year, based on the output from the most recent call of a get_projects command.
        Must be called immediately after the corresponding get_projects or plot_gp call. 
        
        Parameters
        ----------
        save_it : bool, default None
            If True, plot created is saved in .png format. 
        save_name : str or None, default None
            The name plot created should be saved with.
            If None, plot is saved with name "Themes_by_FY".
            Overwrites any pre-existing file of the same name in the save folder.
        
        Returns
        ----------
        None
        """
        
        import sys
        
        #If save_it is not of type boolean, exit
        if type(save_it)!= bool:
           sys.exit("'save_it' must be a boolean.") 
        
        #If the previous command is not the get_projects command, exit
        if ((self.__last_command!="get_projects") and 
            (self.__last_command!="plot_gp") and 
            (self.__last_command!="plot_fy") and 
            (self.__last_command!="save_last")):
           sys.exit("Command not preceded by 'get_projects'.") 
        
        #Record function call in the last_command attribute
        self.__last_command = "plot_fy"
        
        #If the previous get_projects command did not return a valid output, alert user
        if self.__last_output_exist == False:
            print("No output to plot.")
            
        else:
            
            #If save_name is none, set it to a default value
            if save_name==None:
                save_name = "Themes_by_FY"
            else:
                #If save_name is specified and it is not a str, return error
                if type(save_name)!=str:
                    raise TypeError("save_name must be of type 'str'.")
                    
            #Import plt
            from matplotlib import pyplot as plt
            
            #Extract the output_df from the last get_projects command
            output_df = self.__last_output
            
            #Drop duplicates on Project Id
            unique_df = output_df.drop_duplicates(subset="Project Id", keep="first")
            #Count number of Project Ids missing Approval FY data
            missing_FY_count = unique_df["Project Approval FY"].isna().sum()
            #If missing_FY_count > 0, notify user of of the impending exclusion
            if missing_FY_count>0:
                print(f"WARNING! {missing_FY_count} project(s) with missing Approval FY data got excluded from the plot.")
            
            #Group the output_df by GP and count the number of unique project ID
            plot_series = output_df.groupby("Project Approval FY")["Project Id"].nunique()
            
            #Reset the index to transform plot series to dataframe
            plot_df = plot_series.reset_index()
            
            #Sort data
            plot_df.sort_values("Project Approval FY", ascending=True, inplace=True)
            
            #Create plot canvass
            fig, ax = plt.subplots()
            
            #Create horizontal bar chart, using a container to save the chart
            fig_con = ax.bar(plot_df["Project Approval FY"],plot_df["Project Id"])
            
            #Add data label
            ax.bar_label(fig_con)
            
            #Add title to X-Axis
            ax.set_xlabel("Project Approval FY")
            
            #Add title to Y-Axis
            ax.set_ylabel("Number of projects")
            
            #Add chart title
            ax.set_title("Project count, by Project Approval FY")
            
            if save_it:
                print("Plot saved.")
            
            fig.savefig(save_name, bbox_inches='tight') if save_it else plt.show()