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
                  "This typically takes 1-2 minutes. Please wait...")  
            
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
                  "Data source: World Bank Standard Reports." + "\n" +
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
            
            import numpy as np
            
            #Check if any auxiliary argument is provided
            aux_args = (start_FY==None) & (stop_FY==None) & (product_type==None) & (project_status==None) & (include_AF==True)
            
            #USER INPUT VALIDATION
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
                print("WARNING! One or more items in 'sector_codes' is not of the accepted sector hierarchy.")
            
            #---------------------------------------------#
            #If min_pct is not an int, return error
            if type(min_pct)!=int:
                raise TypeError("'min_pct' must be of type 'int'.") 
            #If min_pct is <0 or >100, warn user
            elif min_pct not in range(0,101):
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
                self.__last_output_exist = False
            else: 
                no_of_unique = output_df['Project Id'].nunique()
                print(f"{no_of_unique} unique projects meet the specified criteria.") 
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
            elif any(type(pid)!=str for pid in pid_list):
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
                self.__last_output_exist = False
                print("No data found for specified PID(s).")
            else:
                no_of_unique = output_df['Project Id'].nunique()
                print(f"Data found for {no_of_unique} out of {len(pids)} requested PIDs.")
                self.__last_output = output_df
                self.__last_output_exist = True
                return output_df
    ################################################ 

    def count_sectors(self, pid_list, summarize=False):
        """
        Description
        ----------
        Returns the number of sub-sectors that the specified projects are mapped to.
        Supports variable assignment.
        Data must already be loaded into the Sectors object.
        
        Parameters
        ----------
        pid_list : list
            The ID numbers of projects for which sub-sector counts are to be returned. 
        summarize : bool, default False
            If True, returns frequency count of each sub-sector count.
        
        Returns
        ----------
        DataFrame object
        """
        
        #Record command call in the last_command attribute
        self.__last_command = "count_sectors"
        
        #If summarize is not boolean, return error
        if type(summarize)!=bool:
            raise TypeError("'summarize' must be of type 'bool'.")
        
        #Call the get_sectors command on the pid_list and save output
        temp_data = self.get_sectors(pid_list)
        
        #If get_sectors yielded an empty df, report it and exit
        if self.__last_output_exist == False:
            print("")
        
        #Otherwise, begin computation sequence for count_sect
        else:
            #Import the required libraries
            from collections import defaultdict
            import numpy as np
            import pandas as pd
            
            #--------------------------------------------#
            
            #Replace sector name with sector code if sector name is missing, creating a new column
            temp_data['Sector Name_'] = np.where(temp_data['Sector Long Name'].isnull(),
                                                    temp_data['Sector Code'],
                                                    temp_data['Sector Long Name'])
            
            #--------------------------------------------#
            
            #Create a dictionary that stores all the sub-sector codes for each PID 
            PIDs_n_sectors = defaultdict(dict)
            
            #for each pid, sector code and sector percentage combination in the data
            for pid,sector,pct in zip(temp_data['Project Id'],temp_data['Sector Name_'],temp_data['Sector Percentage']):
                #Update the PIDs_n_sectors using PID as key and dict of sector code and sector percentage as value
                PIDs_n_sectors[pid].update({sector:pct})
                
            #--------------------------------------------#
            
            #Create a dictionary that stores the number of sub-sectors each PID is mapped to
            sector_counts = defaultdict(int) 
            
            #for each PID key-value pairs in the PIDs_n_sectors dict
            for pid,sub_sectors in PIDs_n_sectors.items():
                
                #collect all the sub-sector codes for that pid
                all_sub_sectors = set(sub_sectors.keys())
                
                #discard null values from the set
                all_sub_sectors.discard(np.nan)
                
                #Count the number of unique sector codes mapped to that PID and assign it to the pid key
                sector_counts[pid] = len(all_sub_sectors)
                
            #--------------------------------------------#
            
            #Convert the output to a dataframe
            sector_counts = dict(sector_counts)
            temp_df = pd.DataFrame.from_dict(sector_counts, 
                                                   orient="index",
                                                   columns=["Sector Counts"])
            output_df = temp_df.reset_index()
            output_df.rename(columns={"index":"PID"}, inplace=True)
            
            #--------------------------------------------#
            
            #If summarize is set to True
            if summarize:
                #Create pivot table of the number of PIDs that have each sector count
                plot_df = output_df.pivot_table(values="PID",index="Sector Counts",aggfunc="count")
                plot_df.reset_index(inplace=True)
                plot_df.rename(columns={"PID":"Frequency"}, inplace=True)
                plot_df.sort_values(['Sector Counts'], ascending=True, inplace=True)
                
                from matplotlib import pyplot as plt 
                
                #Create plot canvass
                fig, ax = plt.subplots()
                
                #Create horizontal bar chart, using a container to save the chart
                fig_con = ax.barh(plot_df['Sector Counts'],plot_df["Frequency"])
                
                #Add data label
                ax.bar_label(fig_con)
                
                # make the x ticks integers, not floats
                x_label_int = []
                locs, labels = plt.xticks()
                for each in locs:
                    x_label_int.append(int(each))
                plt.xticks(x_label_int)
                
                #Add title to X-Axis
                ax.set_xlabel("Project count")
                
                #Add title to Y-Axis
                ax.set_ylabel("Number of sectors projects are mapped to")
                ax.invert_yaxis()
                
                #Add chart title
                ax.set_title("Project counts, by Sector Counts")
                
                plt.show()
                    
            self.__last_output = output_df
            self.__last_output_exist = True
            #Re-record last command to overwrite that recorded by the get_sectors call
            self.__last_command = "count_sectors"
            return output_df
            
    ################################################
    
    def main_sector(self, pid_list, threshold=None, summarize=False):
        """
        Description
        ----------
        Returns the main sector that each project in the specified list is mapped to.
        Supports variable assignment.
        Data must already be loaded into the Sectors object.
        
        Parameters
        ----------
        pid_list : list
            The ID numbers of projects for which sector codes are to be returned. 
        threshold : int or None, default None
            If None, main sector is the sector that accounts for the largest percentage of a project. If multiple sectors account for the largest share of the project, the main sub-sector for that project is ambiguous.
            If int, the main sub-sector is the sub-sector that accounts for a percentage value greater than or equal to the threshold value.
            For example, if threshold = 70, any sub-sector that account at least 70% of a project is its main sub-sector. If no sub-sector accounts for up to 70%, the main sub-sector is ambiguous.
            CAUTION: if threshold is 50 or less, multiple sub-sectors may account for a percentage greater than or equal to threshold value, but only one will be returned.
        summarize : bool, default False
            If True, returns frequency counts of each main sub-sector.
        
        Returns
        ----------
        DataFrame object
        """
        
        #Record command call in the last_command attribute
        self.__last_command = "main_sector"
    
        #--------------------------------------------#
    
        #If summarize is not a boolean, return error
        if type(summarize)!=bool:
            raise TypeError("'summarize' must be of type 'bool'.")
            
        #If Threshold is specified, validate the specified value
        if threshold!=None:
            #If threshold is not an integer, return error
            if type(threshold)!=int:
                raise TypeError("'threshold' must be of type 'int'.")
            #If threshold is an integer but not in range 0-101
            elif not threshold in range(0,101):
                raise ValueError("'threshold' is outside expected range of 0 to 100.")
            #If threshold is 50 or less, warn user of potential result unreability
            elif threshold<51:
                print(f"WARNING! A 'threshold' value of {threshold}% may give rise to multiple main sectors, but only one will be returned.")
            
        #--------------------------------------------#
        
        #Call the get_sectors command on the pid_list and save output
        temp_data = self.get_sectors(pid_list)
        
        #If get_sectors yielded an empty df, report it and exit
        if self.__last_output_exist == False:
            print("")
        #Otherwise, begin computation sequence for main sector
        else:
            
            #Import the required libraries
            from collections import defaultdict
            import numpy as np
            import pandas as pd
            
            #--------------------------------------------#
            
            #Replace sector name with sector code if sector name is missing, creating a new column
            temp_data['Sector Name_'] = np.where(temp_data['Sector Long Name'].isnull(),
                                                    temp_data['Sector Code'],
                                                    temp_data['Sector Long Name'])
            
            #--------------------------------------------#
            
            #Create a dictionary that stores the sub-sectors for each PID 
            PIDs_n_sectors = defaultdict(dict)
            
            #For each pid, sector code and sector percentage combination in the data
            for pid,sector,pct in zip(temp_data['Project Id'],temp_data['Sector Name_'],temp_data['Sector Percentage']):
                #Update the PIDs_n_sectors using PID as key, and sector code and sector percentage as a dict
                PIDs_n_sectors[pid].update({sector:pct})
                
            #--------------------------------------------#
            
            #Create a dictionary that stores the dominant sub-sector for each PID
            dom_sector = defaultdict(dict)
        
            #If user wants the sector with the highest percentage, whatever that percentage is
            if threshold==None:
                
                #For each PID and its sub-sector dict in PIDs_n_sectors
                for pid,sub_sectors in PIDs_n_sectors.items():
                    
                    #unpack the sub-sector dict into sector names and pct
                    names,pct_values = (list(sub_sectors.keys()),list(sub_sectors.values()))
                    
                    #Get the maximum sector percentage
                    max_pct = max(pct_values)
                    
                    #Determine if there are multiple max sector percentages
                    weak_max = True if pct_values.count(max_pct)>1 else False
                    
                    #If max is a weak max, indicate that there is no dominant sector
                    if weak_max:
                        dom_dict = {"AMBIGUOUS":np.nan}
                    #else, collect the dominanct sector and pct in a dict
                    else:
                        max_pct_idx = pct_values.index(max_pct)
                        dom_dict = {names[max_pct_idx]:max_pct}
                    
                    #Assign dom_dict to that pid
                    dom_sector[pid] = dom_dict
            
            #If user defines main sector as sector whose percentage is higher than a given threshold
            else:
                
                #For each PID and its sub-sector dict in PIDs_n_sectors
                for pid,sub_sectors in PIDs_n_sectors.items():
                    
                    #Initially designate each pid as a multi-sector project
                    dom_sector[pid] = {"AMBIGUOUS":np.nan}
                    
                    #Then, unpack the sub-sector dict into sector names and pct
                    names,pct_values = (list(sub_sectors.keys()),list(sub_sectors.values()))
                    
                    #If project is mapped to only one sector, then invariably that sector code is the dominant one for that project
                    if len(names)==1:
                        #create dict for the dominant sector
                        dom_dict = {names[0]:pct_values[0]}
                        #Re-assign dom_dict to that project
                        dom_sector[pid]=dom_dict
                    
                    #If project is mapped to more than one sector code
                    else:
                        
                        #for each key-value pair in the sub-sector dict
                        for k,v in sub_sectors.items():
                            
                            #If v >= threshold
                            if v>=threshold:
                                #create dict for the dominant sector
                                dom_dict = {k:v}
                                #Re-assign the dict to that project
                                dom_sector[pid]=dom_dict
                            #continue to the next iteration in loop
                            continue
                    continue
                            
            #Convert dom_sector to regular dict
            dom_sector_dict = dict(dom_sector)
            
            #--------------------------------------------#
            
            #Now, convert the nested dom_sector_dict to a pandas dataframe
            pids = []
            dominant_sectors = []
            
            for pid,d in dom_sector_dict.items():
                pids.append(pid)
                dominant_sectors.append(pd.DataFrame.from_dict(d, orient='index'))
        
            output_df = pd.concat(dominant_sectors, keys=pids)
            output_df.reset_index(inplace=True)
            output_df.rename(columns={'level_0':'PID',
                                      'level_1':'Main Sector', 
                                      0:'Sector Percentage'}, 
                             inplace=True)
            
            #--------------------------------------------#
            
            #If summarize is set to True
            if summarize:
                #Create pivot table of the number of PIDs that have each sector count
                plot_df = output_df.pivot_table(values="PID",index="Main Sector",aggfunc="count")
                plot_df.rename(columns={"PID":"Frequency",}, inplace=True)
                plot_df.reset_index(inplace=True)
                plot_df.sort_values(['Frequency'], ascending=False, inplace=True)
                
                from matplotlib import pyplot as plt
                
                #Create plot canvass
                if plot_df["Main Sector"].nunique() in range(20,40):
                    fig, ax = plt.subplots(figsize=(7,7))
                elif plot_df["Main Sector"].nunique()>40:
                    fig, ax = plt.subplots(figsize=(10,10))
                else:
                    fig, ax = plt.subplots()
                
                #Create horizontal bar chart, using a container to save the chart
                fig_con = ax.barh(plot_df['Main Sector'], plot_df["Frequency"])
                
                #Add data label
                ax.bar_label(fig_con)
                
                # make the x ticks integers, not floats
                x_label_int = []
                locs, labels = plt.xticks()
                for each in locs:
                    x_label_int.append(int(each))
                plt.xticks(x_label_int)
                
                #Add title to X-Axis
                ax.set_xlabel("Project count")
                
                ax.invert_yaxis()
                
                #Add chart title
                ax.set_title("Project count, by Main Sectors")
                
                plt.show()
                    
            self.__last_output = output_df
            self.__last_output_exist = True
            #Re-record last command to overwrite that recorded by the get_sectors call
            self.__last_command = "main_sector"
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
        
        try:
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
                #Extract the output df
                output_df = self.__last_output 
                #Save it to save_name
                output_df.to_excel(save_name,index=False)
                print("Data extract saved.")
        except PermissionError:
            print("ERROR! Save was unsuccessful. A file with the same name is currently open.")
    
    ################################################   

    def plot_last(self, plot_by="sectors", save_name=None):
        
        """
        Description
        ----------
        Plots the output from the most recent call of a get_projects or get_sectors command.
        Output can be saved in .png format.
        
        Parameters
        ----------
        plot_by : str, default "sectors"
            The grouping variable over which project counts will be plotted. 
            Default value is "sectors", but other acceptable values are "GP", "FY", "Status", "region" and "Instrument".
        save_name : str or None, default None
            If str, plot created will be saved as a .png file with the string value as the file name.
            Plot will be saved in the current working directory, and will overwrite any pre-existing file of the same name in the save folder.
            If None, plot created will not be saved locally.
        
        Returns
        ----------
        None
        """
            
        #If command not preceeded by get_projects or get_sectors, alert user
        if ((self.__last_command!= "get_projects") and 
              (self.__last_command!= "get_sectors") and
              (self.__last_command!="plot_last")):
            print("Command must be preceded by 'get_projects' or 'get_sectors'.")
            
        #If the previous command did not return a valid output, alert user
        elif self.__last_output_exist == False:
            print("No output to plot.")
        
        #Else, begin data plotting sequence
        else:
            
            #If save_name is specified and it's not a string, return error
            if save_name!=None and type(save_name)!=str:
                raise TypeError("'save_name' must be of type 'str' or None.")
            
            #Extract the output_df from the last command
            output_df = self.__last_output 
            
            #import required libraries
            import numpy as np
            from matplotlib import pyplot as plt
            
            #Replace sector name with sector code if sector name is missing, creating a new column
            output_df['Sector Name_'] = np.where(output_df['Sector Long Name'].isnull(),
                                                    output_df['Sector Code'],
                                                    output_df['Sector Long Name'])
            
            #Convert the plot_by input to variable in the output_df data
            plot_by_dict = {"sectors":"Sector Name_",
                           "gp":"Lead GP/Global Themes",
                           "fy":"Project Approval FY",
                           "region":"Region Name",
                           "instrument":"Lending Instrument Long Name",
                           "status":"Project Status Name"}
            
            #Convert plot_by input to lower. If operation fails, report that it is not one of the acceptable options.  
            try:
                plot_by_var = plot_by_dict[plot_by.lower()]
            except:
                raise ValueError("'plot_by' value is unrecognized. Acceptable values are 'sectors', 'GP', 'FY', STatus', 'Region', and 'Instrument'.")
            
            if plot_by_var not in (output_df.columns):
                raise KeyError("'plot_by' value missing in output from previous command. Consider setting the 'show_meta' argument in previous command to True.")
            
            #Count the number of projects missing data for the plot_by_input
            missing_df = output_df[output_df[plot_by_var].isna()].copy()
            missing_count = missing_df["Project Id"].nunique()
            
            #If missing_FY_count > 0, notify user of of the impending exclusion
            if missing_count>0:
                print(f"WARNING! {missing_count} project(s) with missing values for {plot_by_var} got excluded from the plot.")
                
            #Create pivot table of project count by plot_by_var
            plot_series = output_df.groupby(plot_by_var)["Project Id"].nunique()
            
            #Reset the index to transform plot series to dataframe
            plot_df = plot_series.reset_index()
            
            #Sort data accordingly
            if plot_by_var == "Project Approval FY":
                plot_df.sort_values(plot_by_var, ascending=False, inplace=True)
            else:
                plot_df.sort_values("Project Id", ascending=True, inplace=True)
                
            #Create plot canvass
            if plot_df[plot_by_var].nunique() in range(20,40):
                fig, ax = plt.subplots(figsize=(7,7))
            elif plot_df[plot_by_var].nunique()>40:
                fig, ax = plt.subplots(figsize=(10,10))
            else:
                fig, ax = plt.subplots()
            
            #Create horizontal bar chart, using a container to save the chart
            fig_con = ax.barh(plot_df[plot_by_var],plot_df["Project Id"])
            
            #Add data label
            ax.bar_label(fig_con)
            
            # make the x ticks integers, not floats
            x_label_int = []
            locs, labels = plt.xticks()
            for each in locs:
                x_label_int.append(int(each))
            plt.xticks(x_label_int)
            
            if plot_by_var == "Project Approval FY":
                ax.invert_yaxis()
                
            #Add title to X-Axis
            ax.set_xlabel("Project count")
            
            #Add chart title
            ax.set_title(f"Project count, by {plot_by_var}")
            
            fig.savefig(save_name, bbox_inches='tight') if save_name!=None else plt.show()
            
            if save_name!=None:
                print("Plot saved.")
            
            #Record command in the last_command attribute
            self.__last_command = "plot_last"
    
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
                  "This typically takes 2-4 minutes. Please wait...")  
            
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
                  "Data source: World Bank Standard Reports." + "\n" +
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
            #Filter the data based on output_rows and output_cols, and reset index
            output_df = temp_data.loc[output_rows,output_cols]
            output_df.reset_index(drop=True,inplace=True)
            
            #If user only wants to see the specified theme codes rather than all theme codes the matching projects are mapped to
            if show_all==False:
                #filter the data to show only the selected theme codes at cut-off
                output_df = output_df.loc[output_df['Theme Code'].isin(theme_codes) & 
                                           (output_df['Theme Percentage']>=min_pct), :].copy()
                output_df.reset_index(drop=True,inplace=True)
            
            #--------------------------------------------#
            #Check if output df is empty and alert user accordingly
            if output_df.empty:
                print("No projects meet the specified criteria.")
            else: 
                no_of_unique = output_df['Project Id'].nunique()
                print(f"{no_of_unique} unique projects meet the specified criteria..") 
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
        theme_level : list or None, default None
            If list, returns only the theme codes matching the specified level(s).
            Acceptable values are:
                1: for the highest theme level,
                2: for middle theme level,
                3: for lowest theme level, or any combination of these. 
            If None, returns all eligible themes, regardless of theme level.
        show_meta : bool, default False
            If True, returns additional project-level meta data.
        
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
                print(f"Data found for {no_of_unique} out of {len(pid_list)} requested projects.")
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
        try:
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
                #Extract the output df
                output_df = self.__last_output 
                #Save it to save_name
                output_df.to_excel(save_name,index=False)
                print("Output saved.")
        except PermissionError:
            print("ERROR! Save was unsuccessful. A file with the same name is currently open.")
    ################################################           
            
    def plot_last(self, plot_by="themes", save_name=None):
        
        """
        Description
        ----------
        Plots the output from the most recent call of a get_projects or get_sectors command.
        Output can be saved in .png format.
        
        Parameters
        ----------
        plot_by : str, default "themes"
            The grouping variable over which project counts will be plotted. 
            Default value is "sectors", but other acceptable values are "GP", "FY", "Status", "region" and "Instrument".
        save_name : str or None, default None
            If str, plot created will be saved as a .png file with the string value as the file name.
            Plot will be saved in the current working directory, and will overwrite any pre-existing file of the same name in the save folder.
            If None, plot created will not be saved locally.
        
        Returns
        ----------
        None
        """
        
        #If command not preceeded by get_projects or get_sectors, alert user
        if ((self.__last_command!= "get_projects") and 
              (self.__last_command!= "get_themes") and
              (self.__last_command!="plot_last")):
            print("Command must be preceded by 'get_projects' or 'get_sectors'.")
            
        #If the previous command did not return a valid output, alert user
        elif self.__last_output_exist == False:
            print("No output to plot.")
        
        #Else, begin data plotting sequence
        else:
            
            #If save_name is specified and it's not a string, return error
            if save_name!=None and type(save_name)!=str:
                raise TypeError("'save_name' must be of type 'str' or None.")
            
            #Extract the output_df from the last command
            output_df = self.__last_output 
            
            #import required libraries
            from matplotlib import pyplot as plt
            
            #Convert the plot_by input to variable in the output_df data
            plot_by_dict = {"themes":"Theme Name",
                           "gp":"Lead GP/Global Themes",
                           "fy":"Project Approval FY",
                           "region":"Region Name",
                           "instrument":"Lending Instrument Long Name",
                           "status":"Project Status Name"}
            
            #Convert plot_by input to lower. If operation fails, report that it is not one of the acceptable options.  
            try:
                plot_by_var = plot_by_dict[plot_by.lower()]
            except:
                raise ValueError("'plot_by' value is unrecognized. Acceptable values are 'themes', 'GP', 'FY', STatus', 'Region', and 'Instrument'.")
            
            if plot_by_var not in (output_df.columns):
                raise KeyError("'plot_by' value is missing in output from previous command. Consider setting the 'show_meta' argument in previous command to True.")
            
            #Count the number of projects missing data for the plot_by_input
            missing_df = output_df[output_df[plot_by_var].isna()].copy()
            missing_count = missing_df["Project Id"].nunique()
            
            #If missing_FY_count > 0, notify user of of the impending exclusion
            if missing_count>0:
                print(f"WARNING! {missing_count} project(s) with missing values for {plot_by_var} got excluded from the plot.")
                
            #Create pivot table of project count by plot_by_var
            output_df=output_df.loc[output_df['Theme Level']==3].copy()
            plot_series = output_df.groupby(plot_by_var)["Project Id"].nunique()
            
            #Reset the index to transform plot series to dataframe
            plot_df = plot_series.reset_index()
            
            #Sort data accordingly
            if plot_by_var == "Project Approval FY":
                plot_df.sort_values(plot_by_var, ascending=False, inplace=True)
            else:
                plot_df.sort_values("Project Id", ascending=True, inplace=True)
                
            #Create plot canvass
            if plot_df[plot_by_var].nunique() in range(20,41):
                fig, ax = plt.subplots(figsize=(7,7))
            elif plot_df[plot_by_var].nunique() in range(40,61):
                fig, ax = plt.subplots(figsize=(10,10))
            elif plot_df[plot_by_var].nunique()>60:
                fig, ax = plt.subplots(figsize=(15,15))
            else:
                fig, ax = plt.subplots()
            
            #Create horizontal bar chart, using a container to save the chart
            fig_con = ax.barh(plot_df[plot_by_var],plot_df["Project Id"])
            
            #Add data label
            ax.bar_label(fig_con)
            
            # make the x ticks integers, not floats
            x_label_int = []
            locs, labels = plt.xticks()
            for each in locs:
                x_label_int.append(int(each))
            plt.xticks(x_label_int)
            
            if plot_by_var == "Project Approval FY":
                ax.invert_yaxis()
                
            #Add title to X-Axis
            ax.set_xlabel("Project count")
            
            #Add chart title
            ax.set_title(f"Project count, by {plot_by_var}")
            
            fig.savefig(save_name, bbox_inches='tight') if save_name!=None else plt.show()
            
            if save_name!=None:
                print("Plot saved.")
            
            #Record command in the last_command attribute
            self.__last_command = "plot_last"