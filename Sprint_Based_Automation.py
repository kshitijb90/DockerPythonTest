# Below packages will be required for pulling "In Sprint Automation Report" Report

import xlsxwriter
import xlrd
from pyral import Rally, rallyWorkset, RallyRESTAPIError
import sys

safeApiKey = sys.argv[1]
print(safeApiKey)
print("Hello I am in program sprint based automation")
def main(args):
    

    # Reading the input excel File & storing rows (Scrum Teams) in a list named as scrum_TeamName.
    scrum_TeamName = []
    #RallyPackageLocation = str(sys.argv[1])
    # RallyPackageLocation = 'C:/Users/ssriva41/Rally_ProjectReports_New'

    #location_of_InputFile = f'{RallyPackageLocation}/Input Data'
    location_of_InputFile = 'Input Data'
    #location_of_InputFile = f'{RallyPackageLocation}/Input_Data.xlsx'
    workbook = xlrd.open_workbook(location_of_InputFile)
    worksheet1 = workbook.sheet_by_name("Projects And Iterations")

    # Reading CA Agile Central Server & Confidential API Key *** User Specific ***
    server = 'rally1.rallydev.com'
    safeApiKey = sys.argv[1]
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
        iteration_list.append(worksheet1.cell_value(Iteration, 1))
    while "" in iteration_list:
        iteration_list.remove("")
    iteration_list.pop(0)
    print("")
    print(">>>> ** Sprint Based Automation Report will be pulled for below 'Projects' & respective 'Iterations' **")
    print("")
    print("  >> Sprint Based Automation Report will be pulled for ",len(scrum_TeamName),"Agile Team/Teams")
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
    #location_of_OutputFile = f'{RallyPackageLocation}/Sprint_Based_Automation.xlsx'
    location_of_OutputFile = 'Sprint_Based_Automation.xlsx'
    workbook = xlsxwriter.Workbook(location_of_OutputFile)

    # Creating FirstTab with Name **Capacity_vs_Estimate**
    worksheet1 = workbook.add_worksheet('Sprint_Based__Automation')
    worksheet1.set_column('A:Z', 25)
    header_format = workbook.add_format({'bold': True,'font_color': 'black','align': 'center','valign': 'vcenter',
                                            'bg_color': 'yellow','border': True,'font_size': 12})
    row_format = workbook.add_format(
        {'bold': False, 'font_color': 'blue', 'align': 'center', 'valign': 'vcenter', 'border': True, 'font_size': 11})
    # Header columns
    temp1_columns_name = 'Project Iteration_Name Total_TestCases Automated_TestCases Automation_Percent'

    worksheet1_header_columns = temp1_columns_name.split()
    for i in range(len(worksheet1_header_columns)):
        worksheet1.write(0, i, worksheet1_header_columns[i], header_format)

    print("")
    print("  >> Sprintvise Automation% Report will be created at below location")
    print("    ", location_of_OutputFile)
    print("")
    i = 1

    # declaring lists which will be required.
    TeamCapacity = []
    row_output = []
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
                Automated_TestCases = 0
                Total_TestCases = 0
                query_Iteration_Name = f'Iteration.Name = "{iteration}"'
                response_UserStory = rally.get('HierarchicalRequirement', fetch=True, projectScope=False,
                                               query= query_Iteration_Name)
                for userstory in response_UserStory:
                    # print("**********", userstory.FormattedID, "*****************")
                    query_UserStory_Name = f'WorkProduct.FormattedID = "{userstory.FormattedID}"'
                    response_TestCases = rally.get('TestCase', fetch=True, projectScope=False,
                                                   query=query_UserStory_Name)
                    for testcase in response_TestCases:
                        # print(testcase.FormattedID)
                        # print(testcase.Method)
                        Total_TestCases = Total_TestCases +1
                        if testcase.Method == 'Automated':
                            Automated_TestCases = Automated_TestCases +1
                        else:
                            continue

                row_output.append(project_names)
                row_output.append(iteration)
                row_output.append(Total_TestCases)
                row_output.append(Automated_TestCases)
                if Total_TestCases != 0:
                    Percent_of_Automation = int((Automated_TestCases)/int(Total_TestCases)*100)
                    Automation_Percent = f'{round(Percent_of_Automation,2)}%'
                    row_output.append(Automation_Percent)
                else:
                    Automation_Percent = 0
                    row_output.append(Automation_Percent)

                # Iteration count in given Project
                iterationCount_In_project = iterationCount_In_project +1

                print("        ",iterationCount_In_project,":",row_output)
                x = 0
                for x in range(len(row_output)):
                    worksheet1.write(i, x, row_output[x], row_format)
                i = i +1
                row_output.clear()
            a = a+1
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
        #k = input("Please enter any key to exit & hit ENTER key: ", )


main([])
