# Below packages will be required for pulling "Sprint vise Velocity Report"
import xlsxwriter
import xlrd
from pyral import Rally, rallyWorkset, RallyRESTAPIError
import sys

# Reading the input excel File & storing rows (Scrum Teams) in a list named as scrum_TeamName.
scrum_TeamName = []
# RallyPackageLocation = sys.argv[1]
RallyPackageLocation = 'C:/Users/ssriva41/Rally_ProjectReports_New'

location_of_InputFile = f'{RallyPackageLocation}/Input_ScrumTeams.xlsm'
workbook = xlrd.open_workbook(location_of_InputFile)
worksheet1 = workbook.sheet_by_name("Sheet1ScrumTeams")
worksheet2 = workbook.sheet_by_name("UserLevel_Information")

# Reading CA Agile Central Server & Confidential API Key *** User Specific ***
server = 'rally1.rallydev.com'
safeApiKey = worksheet2.cell_value(1, 1)
workspace = 'UHG'

# Reading input SCRUM Teams and storing them in list named scrum_TeamName.
# Also removing top heading & unnecessary blank rows
total_ScrumTeams = worksheet1.nrows
for team in range(total_ScrumTeams):
    scrum_TeamName.append(worksheet1.cell_value(team, 0))
scrum_TeamName.pop(0)
# removing unwanted BLANK rows in case.
while "" in scrum_TeamName:
    scrum_TeamName.remove("")

# Reading input **Iteration Names** and storing them in list named iteration_list.
# Also removing top heading & unnecessary blank rows
iteration_list = []
feature_inputs = worksheet1.nrows
for Iteration in range(feature_inputs):
    iteration_list.append(worksheet1.cell_value(Iteration, 3))
while "" in iteration_list:
    iteration_list.remove("")
iteration_list.pop(0)
print("")
print(">>>> ** In-Sprint Velocity Report will be pulled for below 'Projects' & respective 'Iterations' **")
print("")
print("  >> In-Sprint Velocity Report will be pulled for ",len(scrum_TeamName),"Agile Team/Teams")
teams_count = 1
for teams in scrum_TeamName:
    print("     ", teams_count, "-", teams)
    teams_count = teams_count + 1
print("")
print("  >> For each agile team (quoted above) report will be pulled for below", len(iteration_list), "Iteration/Iterations")
iteration_count = 1
for list_items in iteration_list:
    print("     ", iteration_count,"-", list_items)
    iteration_count = iteration_count +1
print("")

# Creating output excel file with name InSprint_AutomationReport and creating the header for output file
location_of_OutputFile = f'{RallyPackageLocation}/Sprintvise_Velocity.xlsx'
workbook = xlsxwriter.Workbook(location_of_OutputFile)

# Creating FirstTab with Name **Capacity_vs_Estimate**
worksheet1 = workbook.add_worksheet('SprintVise_Velocity')
worksheet1.set_column('A:Z', 25)
header_format = workbook.add_format({'bold': True,'font_color': 'black','align': 'center','valign': 'vcenter',
                                        'bg_color': 'yellow','border': True,'font_size': 12})
row_format = workbook.add_format(
    {'bold': False, 'font_color': 'blue', 'align': 'center', 'valign': 'vcenter', 'border': True, 'font_size': 11})
# Header columns
temp1_columns_name = 'Project Iteration_Name Total_UserStories Accepted_UserStories(Count) Iteration_Velocity	Accepted_UserStories NotAccepted_UserStories NotAccepted_UserStories Lost_Velocity'

worksheet1_header_columns = temp1_columns_name.split()
for i in range(len(worksheet1_header_columns)):
    worksheet1.write(0, i, worksheet1_header_columns[i], header_format)

print("  >> Sprint-vise Velocity Report will be created at below location")
print("    ", location_of_OutputFile)
print("")
i = 1


def main(args):
    # declaring lists which will be required.
    accepted_userstories_and_defects_in_sprint = []
    notaccepted_userstories_and_defects_in_sprint = []
    sum_of_Accepted_PlanEstimate = 0
    sum_of_NotAccepted_PlanEstimate = 0
    row_output = []
    row_output_summary = []
    i = 1
    a = 1

    # Now iterating to all input projects one by one which are saved in list **scrum_TeamName = []**
    # Using try & except for exceptional handling.
    try:
        # Building Rally connection in iterative way for all projects in for given projects in list scrum_TeamName
        for project_names in scrum_TeamName:

            rally = Rally(server=server, apikey=safeApiKey, workspace=workspace, project=project_names)
            print("  >>", a, "- CA Agile Central Connection Established with ", project_names)
            iterationCount_In_project = 0

            # Building Rally connection in iterative way for all iterations in iteration_list
            for iteration in iteration_list:
                query_Iteration_Name = f'Iteration.Name = "{iteration}"'
                response_UserStory = rally.get('HierarchicalRequirement', fetch=True, projectScope=False,
                                               query=query_Iteration_Name)
                for userstory in response_UserStory:
                    if userstory.ScheduleState == 'Accepted':
                        accepted_userstories_and_defects_in_sprint.append(userstory.FormattedID)
                        if userstory.PlanEstimate is not None:
                            Accepted_PlanEstimate = int(userstory.PlanEstimate)
                        else:
                            Accepted_PlanEstimate = 0
                        sum_of_Accepted_PlanEstimate = sum_of_Accepted_PlanEstimate + Accepted_PlanEstimate
                    else:
                        notaccepted_userstories_and_defects_in_sprint.append(userstory.FormattedID)
                        # NotAccepted_PlanEstimate = int(userstory.PlanEstimate)
                        if userstory.PlanEstimate is not None:
                            NotAccepted_PlanEstimate = int(userstory.PlanEstimate)
                        else:
                            NotAccepted_PlanEstimate = 0
                        sum_of_NotAccepted_PlanEstimate = sum_of_Accepted_PlanEstimate + NotAccepted_PlanEstimate

                response_Defects = rally.get('Defect', fetch=True, projectScope=False,
                                                 query=query_Iteration_Name)
                for defect in response_Defects:
                    if defect.ScheduleState == 'Accepted':
                        accepted_userstories_and_defects_in_sprint.append(defect.FormattedID)
                        if defect.PlanEstimate is not None:
                            Accepted_PlanEstimate = int(defect.PlanEstimate)
                        else:
                            Accepted_PlanEstimate = 0
                        sum_of_Accepted_PlanEstimate = sum_of_Accepted_PlanEstimate + Accepted_PlanEstimate
                    else:
                        notaccepted_userstories_and_defects_in_sprint.append(defect.FormattedID)
                        # NotAccepted_PlanEstimate = int(defect.PlanEstimate)
                        if defect.PlanEstimate is not None:
                            NotAccepted_PlanEstimate = int(defect.PlanEstimate)
                        else:
                            NotAccepted_PlanEstimate = 0
                        sum_of_NotAccepted_PlanEstimate = sum_of_Accepted_PlanEstimate + NotAccepted_PlanEstimate


                row_output_summary.append(project_names)
                row_output_summary.append(iteration)
                Total_userstrories_and_Defects_insprint = len(accepted_userstories_and_defects_in_sprint) + len(notaccepted_userstories_and_defects_in_sprint)
                row_output_summary.append(Total_userstrories_and_Defects_insprint)
                row_output_summary.append(len(accepted_userstories_and_defects_in_sprint))
                row_output_summary.append(sum_of_Accepted_PlanEstimate)

                row_output.append(project_names)
                row_output.append(iteration)
                Total_userstrories_insprint = len(accepted_userstories_and_defects_in_sprint)+len(notaccepted_userstories_and_defects_in_sprint)
                row_output.append(Total_userstrories_insprint)
                row_output.append(len(accepted_userstories_and_defects_in_sprint))
                row_output.append(sum_of_Accepted_PlanEstimate)
                Accepted_userstories = '\n'.join(accepted_userstories_and_defects_in_sprint)
                row_output.append(Accepted_userstories)
                NotAccepted_userstories = '\n'.join(notaccepted_userstories_and_defects_in_sprint)
                row_output.append(len(notaccepted_userstories_and_defects_in_sprint))
                row_output.append(NotAccepted_userstories)
                row_output.append(sum_of_NotAccepted_PlanEstimate)
                # print(row_output)

                # Iteration count in given Project
                iterationCount_In_project = iterationCount_In_project + 1
                print("        ", iterationCount_In_project, ":", row_output_summary)
                x = 0
                for x in range(len(row_output)):
                    worksheet1.write(i, x, row_output[x], row_format)
                i = i + 1
                row_output.clear()
                accepted_userstories_and_defects_in_sprint.clear()
                notaccepted_userstories_and_defects_in_sprint.clear()
                sum_of_Accepted_PlanEstimate = 0
                sum_of_NotAccepted_PlanEstimate = 0
                row_output.clear()
                row_output_summary.clear()
            a = a + 1
        workbook.close()
        print("")
        print("!! Report is successfully completed !!")
    except Exception as error:
        print("")
        print("!! Some error occurred !!")
        print("Error message: ", error)
        import traceback
        traceback.print_exc()
    finally:
        workbook.close()
        print("")
        k = input("Please enter any key to exit & hit ENTER key: ", )


main([])