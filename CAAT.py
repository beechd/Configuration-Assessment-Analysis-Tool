import xlsxwriter
import pandas as pd

import sys

light_font = '#FFFFFF'
dark_font = '#000000'

medium_bg = '#A0A0A0'
dark_bg = '#000000'

analysis_sheet = ''

workbook = xlsxwriter.Workbook('APM_Analysis.xlsx')

def getListOfApplications():
	frame = pd.read_excel(analysis_sheet, sheet_name='Analysis', engine='openpyxl').dropna()
	return frame['name'].tolist()

def overallAppStatus(application, tasklist):
	frame = pd.read_excel(analysis_sheet, sheet_name='Analysis', engine='openpyxl')
	frame = frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['name'] == application]
	
	# Overall Assessment
	ranking = 'NA'
	if (appFrame['OverallAssessment'] == 'bronze').any():
		ranking = 'Bronze'
	elif (appFrame['OverallAssessment'] == 'silver').any():
		ranking = 'Silver'
	elif (appFrame['OverallAssessment'] == 'gold').any():
		ranking = 'Gold'
	elif (appFrame['OverallAssessment'] == 'platinum').any():
		ranking = 'Platinum'
	return ranking

def appAgentStatus(application, taskList):
	# Sheet name may have changed to AppAgentsAPM
	frame = pd.read_excel(analysis_sheet, sheet_name='AppAgentsAPM', engine='openpyxl')
	frame = frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Agent Metric Limit
	if (appFrame['metricLimitNotHit'] == False).any():
		taskList[2].append("Application Agent metric limit has been reached")

	# Agent Versions
	if (appFrame['percentAgentsLessThan2YearsOld'] < 50).any():
		taskList[0].append(str(100 - int(appFrame['percentAgentsLessThan2YearsOld'])) + '% of Application Agents are 2+ years old')
	elif (appFrame['percentAgentsLessThan1YearOld'] < 80).any():
		taskList[0].append(str(100 - int(appFrame['percentAgentsLessThan1YearOld'])) + '% of Application Agents are at least 1 year old')

	# Agents reporting data
	if (appFrame['percentAgentsReportingData'] < 100).any():
		taskList[0].append(str(100 - int(appFrame['percentAgentsReportingData'])) + "% of Application Agents aren't reporting data")

	if (appFrame['percentAgentsRunningSameVersion'] < 100).any():
		taskList[0].append('Multiple Application Agent Versions')
		
def machineAgentStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='MachineAgentsAPM', engine='openpyxl')
	frame = frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

def businessTranStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='BusinessTransactionsAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Number of Business Transcations
	if (appFrame['numberOfBTs'] > 200).any():
		taskList[1].append("Reduce amount of Business transactions from " + str(int(appFrame['numberOfBTs'])))

	# % of Business Transactions with load
	if (appFrame['percentBTsWithLoad'] < 90).any():
		taskList[1].append(str(100 - int(appFrame['percentBTsWithLoad'])) + '% of Business Transactions have no load over the last 24 hours')

	# Business Transaction Lockdown
	if (appFrame['btLockdownEnabled'] == False).any():
		taskList[1].append("Business Transaction Lockdown is disabled")

	# Number of Custom Match Rules
	if (appFrame['numberCustomMatchRules'] < 3).any():
		if (appFrame['numberCustomMatchRules'] == 0).any():
			taskList[2].append('No Custom Match Rules')
		else:
			taskList[2].append('Only ' + str(int(appFrame['numberCustomMatchRules'])) + ' Custom Match Rules')

def backendStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='BackendsAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# % of Backends with load
	if (appFrame['percentBackendsWithLoad'] < 75).any():
		taskList[2].append(str(100 - int(appFrame['percentBackendsWithLoad'])) + '% of Backends have no load')

	# Backend limit not hit
	if (appFrame['backendLimitNotHit'] == False).any():
		taskList[2].append('Backend limit has been reached')

	# Number of Custom Backend Rules
	if (appFrame['numberOfCustomBackendRules'] == 0).any():
		taskList[2].append('No Custom Backend Rules')

def overheadStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='OverheadAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Developer Mode Not Enabled for any Business Transaction
	if (appFrame['developerModeNotEnabledForAnyBT'] == False).any():
		taskList[2].append('Development Level monitoring is enabled for a Business Transaction')

	# find-entry-points not enabled
	if (appFrame['findEntryPointsNotEnabled'] == False).any():
		taskList[2].append('Find-entry-points node property is enabled')

	# Aggressive Snapshotting not enabled
	if (appFrame['aggressiveSnapshottingNotEnabled'] == False).any():
		taskList[2].append('Aggressive snapshot collection is enabled')

	# Developer Mode not enabled for an application
	if (appFrame['developerModeNotEnabledForApplication'] == False).any():
		taskList[2].appendt('Development Level monitoring is enabled for an Application')

def serviceEndpointStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='ServiceEndpointsAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Number of Custom Service Endpoint Rules
	if (appFrame['numberOfCustomServiceEndpointRules'] == 0).any():
		taskList[2].append('No Custom Service Endpoint rules')

	# Service Endpoint Limit not hit
	if (appFrame['serviceEndpointLimitNotHit'] == False).any():
		taskList[2].append('Service Endpoint limit has been reached')

	# % of enabled Service Endpoints with load
	if (appFrame['percentServiceEndpointsWithLoadOrDisabled'] < 75).any():
		taskList[2].append(str(100 - int(appFrame['percentServiceEndpointsWithLoadOrDisabled'])) + '% of enabled Service Endpoints have no load')

def errorConfigurationStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='ErrorConfigurationAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Sucess Percentage of Worst Transaction
	if (appFrame['successPercentageOfWorstTransaction'] < 80).any():
		taskList[3].append('Some Business Transactions fail ' + str(100 - int(appFrame['successPercentageOfWorstTransaction'])) + '% of the time')

	# Number of Custom rules
	if (appFrame['numberOfCustomRules'] == 0).any():
		taskList[2].append('No custom error configurations')

def healthRulesAlertingStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='HealthRulesAndAlertingAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Number of Health Rule Violations in last 24 hours
	if (appFrame['numberOfHealthRuleViolationsLast24Hours'] > 10).any():
		taskList[3].append(str(int(appFrame['numberOfHealthRuleViolationsLast24Hours'])) + ' Health Rule Violations in 24 hours')

	# Number of modifications to default Health Rules
	if (appFrame['numberOfDefaultHealthRulesModified'] < 2).any():
		if (appFrame['numberOfDefaultHealthRulesModified'] < 2).any():
			taskList[3].append('No modifications to the default Health Rules')
		else:
			taskList[3].append('Only ' + str(int(appFrame['numberOfDefaultHealthRulesModified'])) + ' modifications to the default Health Rules')

	# Number of actions bound to enabled policies
	if (appFrame['numberOfActionsBoundToEnabledPolicies'] < 1).any():
		taskList[3].append('No actions bound to enabled policies')

	# Number of Custom Health Rules
	if (appFrame['numberOfCustomHealthRules'] < 5).any():
		if (appFrame['numberOfCustomHealthRules'] == 0).any():
			taskList[3].append('No Custom Health Rules')
		else:
			taskList[3].append('Only ' + str(int(appFrame['numberOfCustomHealthRules'])) + ' Custom Health Rules')

def dataCollectorStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='DataCollectorsAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Number of data collector fields configured
	if (appFrame['numberOfDataCollectorFieldsConfigured'] < 5).any():
		if (appFrame['numberOfDataCollectorFieldsConfigured'] == 0).any():
			taskList[2].append('No configured Data Collectors')
		else:
			taskList[2].append('Only ' + str(int(appFrame['numberOfDataCollectorFieldsConfigured'])) + ' configured Data Collectors')

	# Number of data collector fields colleced in snapshots in last 24 hours
	if (appFrame['numberOfDataCollectorFieldsCollectedInSnapshotsLast1Day'] < 5).any():
		if (appFrame['numberOfDataCollectorFieldsCollectedInSnapshotsLast1Day'] == 0).any():
			taskList[2].append('No Data Collector fields collected in APM Snapshots in 24 hours')
		else:
			taskList[2].append('Only ' + str(int(appFrame['numberOfDataCollectorFieldsCollectedInSnapshotsLast1Day'])) + ' Data Collector fields collected in APM Snapshots in 24 hours')

	# Number of data collector fields collect in analytics in last 24 hours
	if (appFrame['numberOfDataCollectorFieldsCollectedInAnalyticsLast1Day'] < 5).any():
		if (appFrame['numberOfDataCollectorFieldsCollectedInAnalyticsLast1Day'] == 0).any():
			taskList[2].append('No Data Collector fields collected in Analytics in 24 hours')
		else:
			taskList[2].append('Only ' + str(int(appFrame['numberOfDataCollectorFieldsCollectedInAnalyticsLast1Day'])) + ' Data Collector fields collected in Analytics in 24 hours')

	# BiQ enabled
	if (appFrame['biqEnabled'] == False).any():
		taskList[2].append('BiQ is disabled')

def apmDashBoardsStatus(application, taskList):
	frame = pd.read_excel(analysis_sheet, sheet_name='DashboardsAPM', engine='openpyxl')
	frame.drop('controller', axis = 1)
	appFrame = frame.loc[frame['application'] == application]

	# Number of custom dashboards
	if (appFrame['numberOfDashboards'] < 5).any():
		if (appFrame['numberOfDashboards'] == 1).any():
			taskList[4].append('Only 1 Custom Dashboard')
		elif (appFrame['numberOfDashboards'] == 0).any():
			taskList[4].append('No Custom Dashboards')
		else:
			taskList[4].append('Only ' + str(int(appFrame['numberOfDashboards'])) + ' Custom Dashboards')

	# % of Custom Dashboards modified in last 6 months
	if (appFrame['percentageOfDashboardsModifiedLast6Months'] < 100).any():
		taskList[4].append(str(100 - int(appFrame['percentageOfDashboardsModifiedLast6Months'])) + '% of Custom Dashboards have not been updated in 6+ months')

	# Number of Custom Dashboards using BiQ
	if (appFrame['numberOfDashboardsUsingBiQ'] == 0).any():
		taskList[4].append('No Custom Dashboards using BiQ')

def performAnalysis(application, taskList):
	overallRanking = overallAppStatus(application, taskList)
	appAgentStatus(application, taskList)
	machineAgentStatus(application, taskList)
	businessTranStatus(application, taskList)
	backendStatus(application, taskList)
	overheadStatus(application, taskList)
	serviceEndpointStatus(application, taskList)
	errorConfigurationStatus(application, taskList)
	healthRulesAlertingStatus(application, taskList)
	dataCollectorStatus(application, taskList)
	apmDashBoardsStatus(application, taskList)
	
	return overallRanking

def buildOutput(applicationData, worksheet):
	stepFormat = workbook.add_format()
	stepFormat.set_align('center')
	row_num = 1
	for application in range(0, len(applicationData)):
		applicationName = applicationData[application][0]
		applicationRank = applicationData[application][1]
		currentApplication = applicationData[application][2]
		generateApplicationHeader(applicationName, applicationRank, worksheet, row_num)
		row_num += 1
		task_num = 1
		for categoryIndex in range(0, len(currentApplication)):
			for taskIndex in range(0, len(currentApplication[categoryIndex])):
				worksheet.write(row_num, 0, applicationName)
				worksheet.write(row_num, 1, task_num, stepFormat)
				worksheet.write(row_num, 3, currentApplication[categoryIndex][taskIndex])
				worksheet.write(row_num, 4, applicationRank)
				worksheet.write(row_num, 5, " ")

				if categoryIndex == 0:
					worksheet.write(row_num, 2, "Agent Installation (APM, Machine, Server, etc.)")
				elif categoryIndex == 1:
					worksheet.write(row_num, 2, "Business Transactions")
				elif categoryIndex == 2:
					worksheet.write(row_num, 2, "Advanced Configurations")
				elif categoryIndex == 3:
					worksheet.write(row_num, 2, "Health Rules & Alerts")
				elif categoryIndex == 4:
					worksheet.write(row_num, 2, "Dashboard")
				worksheet.set_row(row_num, None, None, {'level': 1})
				row_num += 1
				task_num += 1
	worksheet.autofilter('A1:E' + str(row_num))


def generateApplicationHeader(applicationName, applicationRank, worksheet, row_num):
	appHeaderFormat = workbook.add_format()
	appHeaderFormat.set_bold()
	appHeaderFormat.set_font_color(dark_font)
	appHeaderFormat.set_bg_color(medium_bg)

	stepsHeaderFormat = workbook.add_format()
	stepsHeaderFormat.set_bold()
	stepsHeaderFormat.set_font_color(dark_font)
	stepsHeaderFormat.set_bg_color(medium_bg)
	stepsHeaderFormat.set_align('center')

	worksheet.write(row_num, 0, applicationName, appHeaderFormat)
	worksheet.write(row_num, 1, " ", stepsHeaderFormat)
	worksheet.write(row_num, 2, " ", appHeaderFormat)
	worksheet.write(row_num, 3, " ", appHeaderFormat)
	worksheet.write(row_num, 4, applicationRank, appHeaderFormat)
	worksheet.write(row_num, 5, " ", appHeaderFormat)

def generateHeaders():
	worksheet = workbook.add_worksheet('Sheet Number 1')

	headerFormat = workbook.add_format()
	headerFormat.set_bold()
	headerFormat.set_font_color(light_font)
	headerFormat.set_bg_color(dark_bg)

	stepFormat = workbook.add_format()
	stepFormat.set_bold()
	stepFormat.set_font_color(light_font)
	stepFormat.set_bg_color(dark_bg)
	stepFormat.set_align('center')

	worksheet.write('A1', 'Application', headerFormat)
	worksheet.write('B1', 'Steps', stepFormat)
	worksheet.write('C1', 'Activity', headerFormat)
	worksheet.write('D1', 'Task', headerFormat)
	worksheet.write('E1', 'Ranking', headerFormat)
	worksheet.write('F1', 'Target', headerFormat)
	
	worksheet.set_column('A:A', 40)
	worksheet.set_column('B:B', 7)
	worksheet.set_column('C:C', 50)
	worksheet.set_column('D:D', 65)
	worksheet.set_column('E:E', 10)

	worksheet.freeze_panes(1, 0)
	return worksheet

if __name__ == "__main__":
	if len(sys.argv) < 2:
		analysis_sheet = r'DefaultJob-MaturityAssessment-apm.xlsx'
	else:
		analysis_sheet = str(sys.argv[1])
		
	worksheet = generateHeaders()

	applicationNames = getListOfApplications()
	applicationData = []

	counter = 1

	for application in applicationNames:
		taskList = [[], [], [], [], []]
		ranking = performAnalysis(application, taskList)
		if ranking != "Platinum":
			for task in taskList:
				if (len(task) > 0):
					applicationData.append([application, ranking, taskList])
					break
		else:
			applicationData.append([application, ranking, []])
		print('Finished application ' + application + ' (' + str(counter) + ')')
		counter += 1

	buildOutput(applicationData, worksheet)
	workbook.close()
