import smartsheet
from smartsheet import folders
import datetime 
from datetime import date
# =========================================================================================================================
# GET DATE AND CREATE LOG FILE
# =========================================================================================================================
my_current_date = date.today()
my_file_date = str(my_current_date)
my_log_file = "smartsheet-metrics-log-" + my_file_date + ".txt"
my_log = open(my_log_file, 'w')
my_log.write("LOG FILE: {0}".format(my_log_file))
# =========================================================================================================================
# CONNECT TO SMARTSHEET, GET LIST OF WORKSPACES "TOKEN" HAS ACCESS TO
# =========================================================================================================================
smartsheet = smartsheet.Smartsheet(<<insert-your-token-here>>)
action = smartsheet.Workspaces.list_workspaces(page_size=20, page=1)
my_log.write("\n- Successfully Connected to SmartSheet.")
assert isinstance(action.total_pages, object)
pages = action.total_pages
workspaces = action.data
# =========================================================================================================================
# LOOP THROUGH WORKSPACES ID HAS ACCESS TO AND CONNECT TO SS REPORT:
# =========================================================================================================================
# INSTANTIATE DICTIONARY OBJECT FOR PROJECT DATA
my_project_dict = {}
for myworkspace in workspaces:
    assert isinstance(myworkspace, object)
    # FIND THE WORKSPACE
    if myworkspace.id == <<add your SS ID here>>:
        my_log.write("\n- Smartsheet Workspace ID: <add your SS ID here>>")
        # GET THE WORKSPACE NAME
        my_ss_workspace = smartsheet.Workspaces.get_workspace(myworkspace.name)
        assert isinstance(my_ss_workspace, object)
        # GET REPORT: MY ORGS PORTFOLIO OF ALL PROJECTS
        my_report = smartsheet.Reports.get_report(<<insert-report-ID>>, page_size=500, page=1)
        my_report_2 = smartsheet.Reports.get_report(<<insert-report-ID>>, page_size=500, page=2)
        # COUNT THE NUMBER OF ROWS IN REPORT
        my_count = my_report.total_row_count
        my_log.write("\n- Total Rows in Report: {0}".format(my_count))
        assert isinstance(my_report, object)
        # GET COLLECTION LIST OF ROWS
        my_rows = my_report.rows
        my_rows2 = my_report_2.rows
        my_rows.extend(my_rows2)
        # SET UP VARIABLES AND SET COUNTS TO ZERO
        my_prj_count = 0
        my_prj_red_count = 0
        my_prj_green_count = 0
        my_prj_yellow_count = 0
        my_prj_open_count = 0
        my_prj_closed_count = 0
        my_prj_on_hold_count = 0
        my_prj_cancel_count = 0
        my_prj_forecast_count = 0
        my_prj_no_status = 0
        my_prj_wpr_count = 0
        my_prj_dc_count = 0
        my_prj_ip_count = 0
        my_prj_uc_count = 0
        my_prj_network_count = 0
        my_prj_acq_count = 0
        my_prj_bfs_count = 0
        my_prj_gov_count = 0
        my_prj_plat_count = 0
        my_prj_gold_count = 0
        my_prj_brnz_count = 0
        my_prj_no_tier = 0
        # INSTANTIATE DICTIONARY OBJECTS FOR SPECIFIC PROJECT DATA
        my_open_projects = {}
        my_closed_projects = {}
        #my_onhold_projects = {}
        # LOOP THROUGH EACH ROW IN THE REPORT
        for my_row in my_rows:
            #my_log.write("\n- Row: {0}".format(my_row.row_number))
            # GET THE COLUMNS WITHIN EACH CELL
            my_cells = my_row.cells
            # CREATE VARIABLES FOR EACH ROW'S CELL DATA
            my_project = ""
            my_prj_color = ""
            my_prj_start = ""
            my_prj_end = ""
            my_prj_status = ""
            my_prj_func = ""
            my_prj_tier = ""
            # LOOP THROUGH EACH CELL IN THE ROW
            for my_cell in my_cells:
                my_log.write("\n	- Cell: {0}".format(my_cell.column_id) + " : {0}".format(my_cell.value) + " ({0}".format(my_cell.display_value) + ")")
				# FIND AND GET THE CELL VALUE FOR PROJECT STATUS COLOR: GREEN/YELLOW/RED
                if my_cell.column_id == 8073883048798084:
                    my_prj_color = my_cell.value
                elif my_cell.column_id == 8430206722566020:
                    my_prj_tier = my_cell.value
                    if my_prj_tier == "Platinum":
                        my_prj_plat_count += 1
                    elif my_prj_tier == "Gold":
                        my_prj_gold_count += 1
                    elif my_prj_tier == "Bronze":
                        my_prj_brnz_count += 1
                    else:
                        my_prj_no_tier += 1
                        my_prj_tier == "None"
					# END IF TIER
                # FIND AND GET THE CELL VALUE FOR FUNCATIONAL AREA
                elif my_cell.column_id == 1865849272330116:
                    my_prj_func = my_cell.value
                    if my_prj_func == "WPR":
                        my_prj_wpr_count += 1
                    elif my_prj_func == "DC Implementations":
                        my_prj_dc_count += 1
                    elif my_prj_func == "Network Projects":
                        my_prj_network_count += 1
                    elif my_prj_func == "Acquisitions":
                        my_prj_acq_count += 1
                    elif my_prj_func == "BFS":
                        my_prj_bfs_count += 1
                    elif my_prj_func == "Infra Provisioning":
                        my_prj_ip_count += 1
                    elif my_prj_func == "UCV Projects":
                        my_prj_uc_count += 1
                    elif my_prj_func == "DCIE Governance":
                        my_prj_gov_count += 1
                # FIND AND GET THE CELL VALUE PROJECT STATUS VALUE: OPEN/CLOSED/ETC
                elif my_cell.column_id == 2909541357643652:
                    my_prj_status = my_cell.value
                    if my_prj_status == "Open":
                        my_prj_open_count += 1
                    elif my_prj_status == "Closed":
                        my_prj_closed_count += 1
                    elif my_prj_status == "On Hold" or my_prj_status == "On-Hold":
                        my_prj_on_hold_count += 1
                    elif my_prj_status == "Cancelled":
                        my_prj_cancel_count += 1
                    elif my_prj_status == "Forecast":
                        my_prj_forecast_count +=1
                    else:
                        my_prj_no_status += 1
                # FIND AND GET THE CELL VALUE FOR PERCENT COMPLETE
                elif my_cell.column_id == 7296062324008836:
                    my_prj_perc_complete = my_cell.display_value
                # FIND AND GET THE CELL VALUE FOR THE PROJECT NAME & COUNT THE # OF PROJECTS
                elif my_cell.column_id == 7058930737145732:
                    my_prj_count += 1
                    #print ("- Project Counter: ", my_prj_count)
                    if my_cell.value is None:
                        my_project = "Undefined"
                    else:
                        my_project = my_cell.value
				# FIND AND GET THE ASSIGNED PM
                elif my_cell.column_id == 7274651240949636:
                    my_assigned_pm = my_cell.value
				# FIND AND GET THE MANAGER
                elif my_cell.column_id == 5129275918575492:
                    my_manager = my_cell.value
				# FIND AND GET THE CELL VALUE FOR THE PROJECT PLANNED START DATE
                elif my_cell.column_id == 8682026124502916:
                    my_prj_start = my_cell.value
                # FIND AND GET THE CELL VALUE FOR THE PROJECT PLANNED END DATE
                elif my_cell.column_id == 519251799893892:
                    my_prj_end = my_cell.value
                # FIND AND GET THE CELL VALUE FOR THE PROJECT ACTUAL START DATE
                # CHECK IF THERE IS NO VALUE AND USE A DEFAULT DATE
                elif my_cell.column_id == 5022851427264388:
                    if my_cell.value is None:
                        my_prj_actual_start = None
                    else:
                        my_prj_actual_start = my_cell.value
                # FIND AND GET THE CELL VALUE FOR THE PROJECT ACTUAL END DATE
                elif my_cell.column_id == 2771051613579140:
                    if my_cell.value is None:
                        my_prj_actual_end = None
                    else:
                        my_prj_actual_end = my_cell.value
                # GET THE REGION
                elif my_cell.column_id == 1750982217492356:
                    if my_cell.value is None:
                        my_prj_region = "NONE"
                    else:
                        my_prj_region = my_cell.display_value
                    # PUSH ALL THE DATA IN THE VARIOUS DICTIONARIES
                    # PROJECT DICTIONARY
                    my_project_dict[my_project] = [my_prj_color, my_prj_tier, my_prj_func, my_prj_status, my_prj_perc_complete, my_prj_start, my_prj_end, my_prj_actual_start, my_prj_actual_end]
                    # OPEN/CLOSED/ON-HOLD PROJECTS DICTIONARY
                    if my_prj_status == "Open":
                        my_open_projects[my_project] = [my_prj_color, my_prj_tier, my_prj_func, my_prj_status, my_prj_perc_complete, my_prj_start, my_prj_end, my_prj_actual_start, my_prj_actual_end]
                    elif my_prj_status == "Closed":
                        my_closed_projects[my_project] = [my_prj_color, my_prj_tier, my_prj_func, my_prj_status, my_prj_perc_complete, my_prj_start, my_prj_end, my_prj_actual_start, my_prj_actual_end]
					# END IF PRJ STATUS
				# END IF COLUMN-ID FOUND, GRAB VALUES
            # NEXT CELL IN FOR LOOP
		# NEXT ROW IN FOR LOOP
    # END IF WORKSPACE = ONE I WANT
# NEXT WORKSPACE IN FOR LOOP
#quit()
# =========================================================================================================================
# USE THE CURRENT DATE AND DETERMINE WHICH QTR AND FISCAL YEAR WE ARE IN
# =========================================================================================================================
my_fy15_q3_end = datetime.date(2015, 4, 25)
my_fy15_q4_start = datetime.date(2015, 4, 26)
my_fy15_q4_end = datetime.date(2015, 7, 25)
# FY16
my_fy16_q1_start = datetime.date(2015, 7, 26)
my_fy16_q1_end = datetime.date(2015, 10, 24)
my_fy16_q2_start = datetime.date(2015, 10, 25)
my_fy16_q2_end = datetime.date(2016, 1, 23)
my_fy16_q3_start = datetime.date(2016, 1, 24)
my_fy16_q3_end = datetime.date(2016, 4, 30)
my_fy16_q4_start = datetime.date(2016, 5, 1)
my_fy16_q4_end = datetime.date(2016, 7, 30)
# FY17
my_fy17_q1_start = datetime.date(2016, 7, 31)
my_fy17_q1_end = datetime.date(2016, 10, 29)
my_fy17_q2_start = datetime.date(2016, 10, 30)
my_fy17_q2_end = datetime.date(2017, 1, 28)
my_fy17_q3_start = datetime.date(2017, 1, 29)
my_fy17_q3_end = datetime.date(2017, 4, 29)
my_fy17_q4_start = datetime.date(2017, 4, 30)
my_fy17_q4_end = datetime.date(2018, 8, 5)
# FY18
my_fy18_q1_start = datetime.date(2018, 8, 6)
#my_fy18_q1_end = datetime.date(2018, 10, 29)
#my_fy18_q2_start = datetime.date(2018, 10, 30)
#my_fy18_q2_end = datetime.date(2019, 1, 28)
#my_fy18_q3_start = datetime.date(2019, 1, 29)
#my_fy18_q3_end = datetime.date(2019, 4, 29)
#my_fy18_q4_start = datetime.date(2019, 4, 30)
#my_fy18_q4_end = datetime.date(2019, 8, 1)
# =========================================================================================================================
# SET VARIABLES FOR EACH QUARTER FOR PROCESSING LATER
# =========================================================================================================================
if my_current_date > my_fy17_q1_start and my_current_date < my_fy17_q1_end:
    my_current_qtr = "Q1"
    my_current_fy = "FY17"
elif my_current_date > my_fy17_q2_start and my_current_date < my_fy17_q2_end:
    my_current_qtr = "Q2"
    my_current_fy = "FY17"
elif my_current_date > my_fy17_q3_start and my_current_date < my_fy17_q3_end:
    my_current_qtr = "Q3"
    my_current_fy = "FY17"
elif my_current_date > my_fy17_q4_start and my_current_date < my_fy17_q4_end:
    my_current_qtr = "Q4"
    my_current_fy = "FY17"
# END IF
# =========================================================================================================================
# COUNT THE PROJECT STATUS FOR ALL OPEN PROJECTS
# =========================================================================================================================
int_open_wpr = 0
int_open_dc = 0
int_open_acq = 0
int_open_net = 0
int_open_ucv = 0
int_open_bfs = 0
int_open_ip = 0
int_open_gov = 0
int_new_prjs = 0
int_open_plat = 0
int_open_gold = 0
int_open_brnz = 0
int_total_open = 0
for my_open in sorted(my_open_projects):
    int_total_open += 1
    my_prj_qtr = ""
    my_prj_fy = ""
    if my_open_projects[my_open][0] == "Green":
        my_prj_green_count += 1
    elif my_open_projects[my_open][0] == "Yellow":
        my_prj_yellow_count += 1
    elif my_open_projects[my_open][0] == "Red":
        my_prj_red_count += 1
    # END IF
    if my_open_projects[my_open][1] == "Platinum":
        int_open_plat += 1
    elif my_open_projects[my_open][1] == "Gold":
        int_open_gold += 1
    elif my_open_projects[my_open][1] == "Bronze":
        int_open_brnz += 1
	# END IF OPEN TIER COUNT
    if my_open_projects[my_open][2] == "WPR":
        int_open_wpr += 1
    elif my_open_projects[my_open][2] == "DC Implementations":
        int_open_dc += 1
    elif my_open_projects[my_open][2] == "Acquisitions":
        int_open_acq += 1
    elif my_open_projects[my_open][2] == "Network Projects":
        int_open_net += 1
    elif my_open_projects[my_open][2] == "UCV Projects":
        int_open_ucv += 1
    elif my_open_projects[my_open][2] == "BFS":
        int_open_bfs += 1
    elif my_open_projects[my_open][2] == "Infra Provisioning":
        int_open_ip += 1
    elif my_open_projects[my_open][2] == "DCIE Governance":
        int_open_gov += 1
    # END IF
    # GET THE NUMBER OF NEW PROJECTS STARTED THIS QUARTER
    # GET THE PROJECTS ACTUAL START DATE TO DETERMINE IF IT IS A NEW PROJECT
    my_prj_begin_date = my_open_projects[my_open][7]
	# IF PROJECT BEGIN DATE IS NOT BLANK/NONE
    if my_prj_begin_date is not None:
        my_prj_yr, my_prj_mm, my_prj_dd = map(int, my_prj_begin_date.split("-"))
        my_start_date = datetime.date(my_prj_yr, my_prj_mm, my_prj_dd)
        if my_start_date > my_fy15_q4_start and my_start_date < my_fy15_q4_end:
            my_prj_qtr = "Q4"
            my_prj_qtr = "FY15"
        elif my_start_date > my_fy16_q1_start and my_start_date < my_fy16_q1_end:
            my_prj_qtr = "Q1"
            my_prj_fy = "FY16"
        elif my_start_date > my_fy16_q2_start and my_start_date < my_fy16_q2_end:
            my_prj_qtr = "Q2"
            my_prj_fy = "FY16"
        elif my_start_date > my_fy16_q3_start and my_start_date < my_fy16_q3_end:
            my_prj_qtr = "Q3"
            my_prj_fy = "FY16"
        elif my_start_date > my_fy16_q4_start and my_start_date < my_fy16_q4_end:
            my_prj_qtr = "Q4"
            my_prj_fy = "FY16"
        elif my_start_date > my_fy17_q1_start and my_start_date < my_fy17_q1_end:
            my_prj_qtr = "Q1"
            my_prj_fy = "FY17"
        elif my_start_date > my_fy17_q2_start and my_start_date < my_fy17_q2_end:
            my_prj_qtr = "Q2"
            my_prj_fy = "FY17"
        elif my_start_date > my_fy17_q3_start and my_start_date < my_fy17_q3_end:
            my_prj_qtr = "Q3"
            my_prj_fy = "FY17"
        elif my_start_date > my_fy17_q4_start and my_start_date < my_fy17_q4_end:
            my_prj_qtr = "Q4"
            my_prj_fy = "FY17"
        elif my_start_date > my_fy17_q4_end:
            my_prj_qtr = "Q1"
            my_prj_fy = "FY18"
        elif my_start_date < my_fy15_q3_end:
            my_prj_qtr = "Q3"
            my_prj_fy = "FY15"
        else:
            my_prj_qtr = "None"
            my_prj_fy = "None"
        # END IF
    # END IF
    # COUNT THE NUMBER OF NEW PROJECTS BASED ON THE CURRENT QUARTER
    if my_current_qtr == my_prj_qtr and my_prj_fy == my_current_fy:
        int_new_prjs += 1
# NEXT PROJECT IN PROJECT DICTIONARY
# =========================================================================================================================
# GET THE NUMBER OF CLOSED PROJECTS FOR THE CURRENT QUARTER
# =========================================================================================================================
int_closed_prjs_qtr = 0
int_on_time = 0
int_not_on_time = 0
for my_key in sorted(my_closed_projects):
    my_prj_qtr = ""
    my_prj_fy = ""
    # GET THE CLOSED PROJECT'S PLANNED END DATA AND ACTUAL END DATE
    my_prj_end_date = my_closed_projects[my_key][8]
    my_prj_plan_end = my_closed_projects[my_key][6]
    # CONVERT ACTUAL END DATE TO DATATIME FUNCTION
    if my_prj_end_date is not None:
        my_prj_yr, my_prj_mm, my_prj_dd = map(int, my_prj_end_date.split("-"))
        my_end_date = datetime.date(my_prj_yr, my_prj_mm, my_prj_dd)
        if my_end_date > my_fy17_q1_start and my_end_date < my_fy17_q1_end:
            my_prj_qtr = "Q1"
            my_prj_fy = "FY17"
        elif my_end_date > my_fy17_q2_start and my_end_date < my_fy17_q2_end:
            my_prj_qtr = "Q2"
            my_prj_fy = "FY17"
        elif my_end_date > my_fy17_q3_start and my_end_date < my_fy17_q3_end:
            my_prj_qtr = "Q3"
            my_prj_fy = "FY17"
        elif my_end_date > my_fy17_q4_start and my_end_date < my_fy17_q4_end:
            my_prj_qtr = "Q4"
            my_prj_fy = "FY17"
        elif my_end_date > my_fy17_q4_end:
            my_prj_qtr = "Q0"
            my_prj_fy = "FY18"
        elif my_end_date < my_fy17_q1_start:
            my_prj_qtr = "Q1"
            my_prj_fy = "FY16"
        else:
            my_prj_qtr = "None"
            my_prj_fy = "None"
        # END IF
	# END IF
	# COUNT THE # OF CLOSED PROEJCT FOR THE CURRENT QUARTER
    if my_current_qtr == my_prj_qtr and my_prj_fy == my_current_fy:
        int_closed_prjs_qtr += 1
	    # CONVERT PLANNED END DATE TO DATETIME FUNCTION
        if my_prj_plan_end is not None:
            my_pln_end_yr, my_pln_end_mm, my_pln_end_dd = map(int, my_prj_plan_end.split("-"))
            my_pln_end_date = datetime.date(my_pln_end_yr, my_pln_end_mm, my_pln_end_dd)
            if my_pln_end_date == my_end_date or my_pln_end_date > my_end_date:
                int_on_time += 1
            else:
                int_not_on_time += 1
            # END IF ON TIME
        # END IF ON TIME
    # END IF CURRENT QUARTER
	# COUNT THE PROJECT TYPE ALL CLOSED PROJECTS
    int_clsd_wpr = 0
    int_clsd_dc = 0
    int_clsd_acq = 0
    int_clsd_net = 0
    int_clsd_ucv = 0
    int_clsd_bfs = 0
    int_clsd_ip = 0
    int_clsd_gov = 0
    if my_closed_projects[my_key][2] == "WPR":
        int_clsd_wpr += 1
    elif my_closed_projects[my_key][2] == "DC Implementations":
        int_clsd_dc += 1
    elif my_closed_projects[my_key][2] == "Acquisitions":
        int_clsd_acq += 1
    elif my_closed_projects[my_key][2] == "Network Projects":
        int_clsd_net += 1
    elif my_closed_projects[my_key][2] == "UCV Projects":
        int_clsd_ucv += 1
    elif my_closed_projects[my_key][2] == "BFS":
        int_clsd_bfs += 1
    elif my_closed_projects[my_key][2] == "Infra Provisioning":
        int_clsd_ip += 1
    elif my_closed_projects[my_key][2] == "DCIE Governance":
        int_clsd_gov += 1
    # END IF
# NEXT FOR EACH CLOSED PROJECT
# =========================================================================================================================
# CALCULATE PERCENTAGES AND FORMAT RESULTS
# =========================================================================================================================
# PROJECT STATE CALCULATIONS
my_perc_open = (my_prj_open_count/my_prj_count)*100
my_perc_closed = (my_prj_closed_count/my_prj_count)*100
my_perc_onhold = (my_prj_on_hold_count/my_prj_count)*100
my_perc_cancel = (my_prj_cancel_count/my_prj_count)*100
my_perc_forecast = (my_prj_forecast_count/my_prj_count)*100
my_perc_open = format(my_perc_open, '.0f')
my_perc_closed = format(my_perc_closed, '.0f')
my_perc_onhold = format(my_perc_onhold, '.0f')
my_perc_cancel = format(my_perc_cancel, '.0f')
my_perc_forecast = format(my_perc_forecast, '.0f')
# CLOSED PROJECTS: ON TIME DELIVERY CALCULATIONS
my_perc_ontime = (int_on_time/int_closed_prjs_qtr)*100
my_perc_ontime = format(my_perc_ontime, '.0f')
# PROJECT STATUS CALCULATIONS
my_perc_green = (my_prj_green_count/len(my_open_projects))*100
my_perc_yellow = (my_prj_yellow_count /len(my_open_projects))*100
my_perc_red = (my_prj_red_count /len(my_open_projects))*100
my_perc_green = format(my_perc_green, '.0f')
my_perc_yellow = format(my_perc_yellow, '.0f')
my_perc_red = format(my_perc_red, '.0f')
# PROJECT TIER CALCULATIONS
my_perc_plat = (my_prj_plat_count/my_prj_count)*100
my_perc_gold = (my_prj_gold_count/my_prj_count)*100
my_perc_brnz = (my_prj_brnz_count/my_prj_count)*100
my_perc_plat = format(my_perc_plat, '.0f')
my_perc_gold = format(my_perc_gold, '.0f')
my_perc_brnz = format(my_perc_brnz, '.0f')
# PROJECT FUNCIONAL AREA OPEN PROJECT CALCULATIONS
my_perc_dc = (int_open_dc/int_total_open)*100
my_perc_acq = (int_open_acq/int_total_open)*100
my_perc_wpr = (int_open_wpr/int_total_open)*100
my_perc_ip = (int_open_ip/int_total_open)*100
my_perc_bfs = (int_open_bfs/int_total_open)*100
my_perc_uc = (int_open_ucv/int_total_open)*100
my_perc_net = (int_open_net/int_total_open)*100
my_perc_gov = (int_open_gov/int_total_open)*100
my_perc_dc = format(my_perc_dc, '.0f')
my_perc_acq = format(my_perc_acq, '.0f')
my_perc_wpr = format(my_perc_wpr, '.0f')
my_perc_ip = format(my_perc_ip, '.0f')
my_perc_bfs = format(my_perc_bfs, '.0f')
my_perc_uc = format(my_perc_uc, '.0f')
my_perc_net = format(my_perc_net, '.0f')
my_perc_gov = format(my_perc_gov, '.0f')
# =========================================================================================================================
# WRITE TO METRICS SMARTSHEET
# =========================================================================================================================
my_metric_sheet = smartsheet.Sheets.get_sheet(<<insert-SS-ID-here>>)
my_metric_values = [my_file_date, my_current_qtr, my_current_fy, my_prj_count, int_new_prjs, int_closed_prjs_qtr, my_perc_ontime + "%", int_open_plat, int_open_gold, int_open_brnz, my_prj_open_count, my_perc_open + "%", my_prj_green_count, my_perc_green + "%", my_prj_yellow_count, my_perc_yellow + "%", my_prj_red_count, my_perc_red + "%", my_prj_on_hold_count, my_perc_onhold + "%", my_prj_closed_count, my_perc_closed + "%", my_prj_cancel_count, my_perc_cancel + "%", my_prj_forecast_count, my_perc_forecast + "%", int_open_dc, my_perc_dc + "%", int_open_wpr, my_perc_wpr + "%", int_open_ip, my_perc_ip + "%", int_open_acq, my_perc_acq + "%", int_open_net, my_perc_net + "%", int_open_ucv, my_perc_uc + "%", int_open_bfs, my_perc_bfs + "%", int_open_gov, my_perc_gov + "%"]
my_len_list = len(my_metric_values)
my_metric_dict = {}
my_metric_rows_dict = {}
if my_metric_sheet:
    my_log.write("\n\n**Write Metric Values to Smartsheet: DCIE Metrics")
    my_log.write("\n- Metric Sheet ID: " + str(my_metric_sheet.id))
    my_log.write("\n- Sheet Access Level: " + str(my_metric_sheet.access_level))
    my_log.write("\n- Sheet Name: " + str(my_metric_sheet.name))
    # Get paginated list of columns (100 columns per page).
    action = smartsheet.Sheets.get_columns(my_metric_sheet.id)
    pages = action.total_pages
    my_metric_columns = action.data
    my_row_a = smartsheet.models.Row()
    my_row_a.to_top = True
    for my_column in my_metric_columns:
        my_metric_dict[my_column.id] = ""
    # NEXT	
    num = 0
    for my_metric_key in my_metric_dict:
        if num < my_len_list:
            my_row_a.cells.append({'column_id': my_metric_key, 'value': str(my_metric_values[num])})
            num += 1
        # END IF
    # NEXT KEY IN METRIC DIC
	# ADD ROW TO SHEET
    my_action = smartsheet.Sheets.add_rows(<<insert-SS-ID-here>>, [my_row_a])
    if my_action.message == "SUCCESS":
        my_log.write("\n- Success! Metrics smartsheet updated with latest metrics!")
    else:
        my_log.write("\n- Error! Metrics smartsheet failed to update!")
    # END IF
else:
    my_log.write("\n- Error: Could not connect to Smartsheet! No sheet ID returned.")
# END IF
# =========================================================================================================================
# WRITE RESULTS TO LOG FILE
# =========================================================================================================================
my_log.write("\n\n**DCIE Metrics for Current Quarter: " + my_current_qtr + " " + my_current_fy)
my_log.write("\n- Total New Projects in " + my_current_qtr + " " + my_current_fy + " = " + str(int_new_prjs))
my_log.write("\n- Total Closed Projects in " + my_current_qtr + " " + my_current_fy + " = " + str(int_closed_prjs_qtr))
my_log.write("\n- Total Closed Project Delivered on time: " + str(int_on_time) + " (" + my_perc_ontime + "%)")
my_log.write("\n\n**Breakdown of Total DCIE Projects (Open/Closed/Cancelled/On-Hold): " + str(my_prj_count))
my_log.write("\n - Total Open DCIE Projects: " + str(my_prj_open_count) + " (" + str(my_perc_open) +"%)")
my_log.write("\n	> Total Open Green DCIE Projects: " + str(my_prj_green_count) + " (" + str(my_perc_green) + "%)")
my_log.write("\n	> Total Open Yellow DCIE Projects: " + str(my_prj_yellow_count) + " (" + str(my_perc_yellow) + "%)")
my_log.write("\n	> Total Open Red DCIE Projects: " + str(my_prj_red_count) + " (" + str(my_perc_red) + "%)")
my_log.write("\n\n**Breakdown by TIER")
my_log.write("\n- Total Platinum Projects: " + str(my_prj_plat_count) + " (" + str(my_perc_plat) + "%)")
my_log.write("\n- Total Gold Projects: " + str(my_prj_gold_count) + " (" + str(my_perc_gold) + "%)")
my_log.write("\n- Total Bronze Projects: " + str(my_prj_brnz_count) + " (" + str(my_perc_brnz) +"%)")
my_log.write("\n\n**Breakdown by DCIE Functional Area")
my_log.write("\n - Total On-Hold DCIE Projects: " + str(my_prj_on_hold_count) + " (" + str(my_perc_onhold) + "%)")
my_log.write("\n - Total Closed DCIE Projects: " + str(my_prj_closed_count) + " (" + str(my_perc_closed) + "%)")
my_log.write("\n - Total Cancelled DCIE Projects: " + str(my_prj_cancel_count) + " (" + str(my_perc_cancel) + "%)")
my_log.write("\n - Total Forecasted DCIE Projects: " + str(my_prj_forecast_count) + " (" + str(my_perc_forecast) + "%)")
my_log.write("\n - Total Projects with no Status: " +  str(my_prj_no_status))
my_log.write("\n - Total DC Implementations Projects: " + str(my_prj_dc_count) + " (" + str(my_perc_dc) + "%)")
my_log.write("\n - Total WPR Projects: " + str(my_prj_wpr_count) + " (" + str(my_perc_wpr) + "%)")
my_log.write("\n - Total Infra Provisioning Projects: " + str(my_prj_ip_count) + " ("+ str(my_perc_ip) + "%)")
my_log.write("\n - Total Acquisition Projects: " + str(my_prj_acq_count) + " (" + str(my_perc_acq) + "%)")
my_log.write("\n - Total Network Projects: " + str(my_prj_network_count) + " (" + str(my_perc_net) + "%)")
my_log.write("\n - Total UCV Projects: " + str(my_prj_uc_count) + " (" + str(my_perc_uc) + "%)")
my_log.write("\n - Total BFS Projects: " + str(my_prj_bfs_count) + " (" + str(my_perc_bfs) + "%)")
my_log.write("\n - Total Governance Projects: " + str(my_prj_gov_count) + " (" + str(my_perc_gov) + "%)\n")
my_log.close()
quit()
# =========================================================================================================================
