import xlsxwriter
import xlrd
from pyral import Rally, rallyWorkset, RallyRESTAPIError
import sys
import time
import os

pwd = os.getcwd()
print("My current working directory", pwd)
Milestone = sys.argv[1]
safeApiKey = sys.argv[2]
server = 'rally1.rallydev.com'
workspace = 'UHG'
project = 'TC - PBM Services'
reportName = f'{pwd}/{Milestone}_Rally_Timebox_Reconcile_Dashboard.xlsx'
print("Report Summary ---", reportName)
location_of_OutputFile = f'{reportName}'
workbook = xlsxwriter.Workbook(location_of_OutputFile)
tab_MileStoneToFeature_TCs = workbook.add_worksheet('Milestone_Summary')
tab_DashBorad = workbook.add_worksheet('Dashboard')
tab_MilestoneToSR = workbook.add_worksheet('MS vs SRs')
# tab_SpikeUnfinishedSummary = workbook.add_worksheet('Spike Unfinished Summary')
header_format = workbook.add_format({'bold': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
                                     'bg_color': 'yellow', 'border': True, 'font_size': 10})
row_format = workbook.add_format(
    {'bold': False, 'font_color': 'blue', 'align': 'center', 'valign': 'vcenter', 'border': True, 'font_size': 11,
     'text_wrap': True})
tab_MileStoneToFeature_TCs.set_column('A:Z', 22)
tab_DashBorad.set_column('A:Z', 30)
tab_MilestoneToSR.set_column('A:Z', 30)
worksheet1_header_columns = ['Milestone IDs', 'Rally ID','SR','DE & US', 'Test Cases Count']
for i in range(len(worksheet1_header_columns)):
    tab_MileStoneToFeature_TCs.write(0, i, worksheet1_header_columns[i], header_format)
def main(args):
    def report():
        try:
            i = 1
            startMessage = f'    For Milestone {Milestone}, Rally Timebox Reconcile Reported Started'
            print("")
            print(startMessage)
            print("")
            rally = Rally(server=server, apikey=safeApiKey, workspace=workspace, project=project)
            if len(Milestone) == 0:
                print("Incorrect input")
                print("Please enter valid milestone id")
                A = input("Please press any key to exit:",)
                exit()
            else:
                pass
            query_milestone = f'FormattedID = "{Milestone}"'
            response_milestones = rally.get('Milestone', fetch=True, projectScopeDown=True, query=query_milestone)
            rallyMileStoneArtifactsList = []
            for milestone_input in response_milestones:
                rallyMileStoneArtifactsList = milestone_input.Artifacts
            for feature in rallyMileStoneArtifactsList:
                if feature.FormattedID[0:1] == 'F':
                    query_criteria_us = f'Feature.FormattedID = "{feature.FormattedID}"'
                    response_userstoriesInFeature = rally.get('HierarchicalRequirement', fetch=True,
                                                              projectScopeDown=True, query=query_criteria_us)
                    for userstory in response_userstoriesInFeature:
                        userstory_title = userstory.Name.lower()
                        if ("spike" or "unfinished")in userstory_title:
                            print("       spike or unfinished user story", userstory.FormattedID, "-",userstory.Name)
                        else:
                            SR = userstory.Feature.c_RequirementID
                            query_criteria = f'WorkProduct.FormattedID = "{userstory.FormattedID}"'
                            response_TestCases = rally.get('TestCase', fetch=True, projectScopeDown=True,
                                                          query=query_criteria)
                            testCaseCount = 0
                            for testcase in response_TestCases:
                                testCaseCount = testCaseCount + 1
                            j=0
                            tab_MileStoneToFeature_TCs.write(i, j, Milestone)
                            tab_MileStoneToFeature_TCs.write(i, j+1, feature.FormattedID)
                            tab_MileStoneToFeature_TCs.write(i, j + 2, userstory.Feature.c_RequirementID)
                            tab_MileStoneToFeature_TCs.write(i, j + 3, userstory.FormattedID)
                            tab_MileStoneToFeature_TCs.write_number(i, j + 4, int(testCaseCount))
                            i = i+1
                if feature.FormattedID[0:1] == 'D':
                    query_criteria_us = f'FormattedID = "{feature.FormattedID}"'
                    response_defectDetails = rally.get('Defect', fetch=True,
                                                              projectScopeDown=True, query=query_criteria_us)
                    for defect in response_defectDetails:
                        SR = defect.c_RequirementID
                        query_criteria = f'WorkProduct.FormattedID = "{defect.FormattedID}"'
                        response_TestCases = rally.get('TestCase', fetch=True, projectScopeDown=True,
                                                      query=query_criteria)
                        testCaseCount = 0
                        for testcase in response_TestCases:
                            testCaseCount = testCaseCount +1
                        j=0
                        tab_MileStoneToFeature_TCs.write(i, j, Milestone)
                        tab_MileStoneToFeature_TCs.write(i, j+1, feature.FormattedID)
                        tab_MileStoneToFeature_TCs.write(i, j + 2, defect.c_RequirementID)
                        tab_MileStoneToFeature_TCs.write(i, j + 3, defect.FormattedID)
                        tab_MileStoneToFeature_TCs.write_number(i, j + 4, int(testCaseCount))
                        i = i+1
                else:
                    continue
            print("")
            print("     For Milestone -", Milestone ," - Below are counts/stats:      ")
        except Exception as error:
            print("")
            print("!! Some error occurred !!")
            print("Error message: ", error)
            import traceback
            traceback.print_exc()
        finally:
            workbook.close()
            print("")
    report()

main([])
