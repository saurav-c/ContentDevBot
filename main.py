import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import datetime
from slack import WebClient
from slack.errors import SlackApiError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Content Dev Schedule Spreadsheet
SPREADSHEET_ID = #FILL IN
DISCUSSION_RANGE = #FILL IN ex.'Discussion!B3:L36'
NOTE_RANGE = #FILL IN ex.'Notes!B3:K36'
VITAMIN_RANGE = #FILL IN ex.'Vitamins!B3:K36'
RECORDING_RANGE = #FILL IN ex. 'Recorded Discussion!C3:G36'

# CHANGE to relevant semester
SEMESTER_START_DATE = datetime.date(2020, 8, 27)

# CHANGE so it matches the spreadsheets
# Column Indexes
D_SECTION_DATE = 0
D_WEEK = 1
D_TOPIC = 2
D_IMPROVER_1 = 3
D_IMPROVER_2 = 4
D_START_DATE = 5
D_IMPROVE_DATE = 6
D_SHARE_DATE = 7
D_REVIEWER = 8
D_REVIEWER_RECV = 9
D_SHARED = 10
D_RELEASED = 11

N_WEEK = 0
N_TOPIC = 2
N_IMPROVER = 3
N_REVIEWER = 4
N_DUE_DATE = 6
N_RECEIVED = 8

V_WEEK = 0
V_TOPIC = 2
V_IMPROVER = 4
V_REVIEWER = 5
V_DUE_DATE = 7
V_RECEIVED = 8

R_WEEK = 0
R_TOPIC = 1
R_RECORDER = 2
R_RELEASE_DATE = 3
R_RELEASED = 4

# Easy hacky way to get user Slack IDs
slackIDS = {
	# FILL IN
	# 'Oski': '123'
}

# Slack Channel
SLACK_CHANNEL_ID = # FILL IN


def main():
	service = getSheetsService()
	sheet = service.spreadsheets()

	sectionJobs  = getDiscussionJobs(sheet)
	noteJobs = getNoteJobs(sheet)
	vitaminJobs = getVitaminJobs(sheet)
	recordingJobs = getRecordedSectionJobs(sheet)

	sendWeeklyMsg(sectionJobs, noteJobs, vitaminJobs, recordingJobs)

def sendWeeklyMsg(sectionJobs, noteJobs, vitaminJobs, recordingJobs):
	currentWeek = getCurrentWeek()
	msg = '*Week ' + str(currentWeek) + ' Update*\n'
	
	if len(sectionJobs) > 0:
		msg += '\n• Discussion'
		for job in sectionJobs:
			msg += '\n\t- '
			msg += job.toMessageSpecial()
		msg += '\n'

	if len(noteJobs) > 0:
		msg += '\n• Notes'
		for job in noteJobs:
			msg += '\n\t- '
			msg += job.toMessage()
		msg += '\n'

	if len(vitaminJobs) > 0:
		msg += '\n• Vitamins'
		for job in vitaminJobs:
			msg += '\n\t- '
			msg += job.toMessage()
		msg += '\n'

	if len(recordingJobs) > 0:
		msg += '\n• Recorded Discussion'
		for job in recordingJobs:
			msg += '\n\t- '
			msg += job.toMessage()
		msg += '\n'
	
	token = # FILL IN
	client = WebClient(token=token)

	# client = WebClient(token=os.environ['SLACK_API_TOKEN'])
	try:
		response = client.chat_postMessage(
			channel=SLACK_CHANNEL_ID,
			text=msg)

		ts = response['ts']

		# Unpin existing pinned messages
		response = client.pins_list(
			channel=SLACK_CHANNEL_ID)
		
		for item in response['items']:
			msgTS = item['message']['ts']
			client.pins_remove(
				channel=SLACK_CHANNEL_ID,
				timestamp=msgTS)
		
		response = client.pins_add(
			channel=SLACK_CHANNEL_ID,
			timestamp=ts)


	except SlackApiError as e:
		# Eventually send Slack DM stating failure
		print(e)

	#print('Finished sending slack message!')

def getDiscussionJobs(sheet):
	result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=DISCUSSION_RANGE).execute()
	values = result.get('values', [])
	
	jobs = []

	curWeek = getCurrentWeek()
	# Find all weeks with rec. start date this week or next week
	for row in values:
		if len(row) > 4:
			week = int(row[D_WEEK])
			if abs(week - curWeek) <= 2 and row[D_SHARED] == 'FALSE':
				jobs.append(DiscussionJob(week=week, topic=row[D_TOPIC], 
					improvers=[row[D_IMPROVER_1], row[D_IMPROVER_2]], reviewer=row[D_REVIEWER],
					improveBy=row[D_IMPROVE_DATE], shareBy=row[D_SHARE_DATE], priority=curWeek==week))

	return jobs

def getNoteJobs(sheet):
	result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=NOTE_RANGE).execute()
	values = result.get('values', [])
	
	jobs = []

	curWeek = getCurrentWeek()
	for row in values:
		if len(row) > 4:
			week = int(row[N_WEEK])
			if abs(week - curWeek) <= 2 and row[N_RECEIVED] == 'FALSE':
				jobs.append(Job(week=week, topic=row[N_TOPIC], improver=row[N_IMPROVER],
					reviewer=row[N_REVIEWER], dueBy=row[N_DUE_DATE]))

	return jobs

def getVitaminJobs(sheet):
	result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=VITAMIN_RANGE).execute()
	values = result.get('values', [])
	
	jobs = []

	curWeek = getCurrentWeek()
	for row in values:
		if len(row) > 4:
			week = int(row[V_WEEK])
			if abs(week - curWeek) <= 2 and row[V_RECEIVED] == 'FALSE':
				jobs.append(Job(week=week, topic=row[V_TOPIC], improver=row[V_IMPROVER],
					reviewer=row[V_REVIEWER], dueBy=row[V_DUE_DATE]))

	return jobs

def getRecordedSectionJobs(sheet):
	result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RECORDING_RANGE).execute()
	values = result.get('values', [])
	
	jobs = []

	curWeek = getCurrentWeek()
	for row in values:
		if len(row) > 3:
			week = int(row[R_WEEK])
			if week <= curWeek and row[R_RELEASED] == 'FALSE':
				jobs.append(RecordingJob(week=week, topic=row[R_TOPIC],
					recorder=row[R_RECORDER], dueBy=row[R_RELEASE_DATE]))

	return jobs




def getSheetsService():
	creds = None
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)

	service = build('sheets', 'v4', credentials=creds)
	return service

def getCurrentWeek():
	week = int(datetime.datetime.now().strftime("%W"))
	return week - int(SEMESTER_START_DATE.strftime("%W")) + 1

def getSlackIDFromName(name):
	return '<@' + slackIDS[name] + '>'

def getSlackIDFromNames(names):
	return [getSlackIDFromName(n) for n in names]


class DiscussionJob:
	def __init__(self, week, topic, improvers, reviewer, improveBy, shareBy, priority=False):
		self.week = week
		self.topic = topic
		self.improvers = getSlackIDFromNames(improvers)
		self.reviewer = getSlackIDFromName(reviewer)
		self.improveByDate = improveBy
		self.shareByDate = shareBy
		self.priority = priority

	def toMessage(self):
		if self.priority:
			return "*_Week {} - {}_*: {} and {} please share today!".format(self.week, self.topic, 
				self.improvers[0], self.improvers[1])
		return "*_Week {} - {}_*: {} and {} send for review by {}. {} to review. Share with section TAs by {}.".format(
				self.week, self.topic, self.improvers[0], self.improvers[1], self.improveByDate, self.reviewer, self.shareByDate)

	def toMessageSpecial(self):
		return "*_Week {} - {}_*: {} and {} and {} send for review by {}. {} to review. Share with section TAs by {}.".format(
				self.week, self.topic, self.improvers[0], self.improvers[1], getSlackIDFromName('Su Min'), self.improveByDate, self.reviewer, self.shareByDate)	

class Job:
	def __init__(self, week, topic, improver, reviewer, dueBy):
		self.week = week
		self.topic = topic
		self.improver = getSlackIDFromName(improver)
		self.reviewer = getSlackIDFromName(reviewer)
		self.dueBy = dueBy

	def toMessage(self):
		return "*_Week {} - {}_*: {} to improve and {} to review. Share for final review by {}.".format(
			self.week, self.topic, self.improver, self.reviewer, self.dueBy)

class RecordingJob:
	def __init__(self, week, topic, recorder, dueBy):
		self.week = week
		self.topic = topic
		self.recorder = getSlackIDFromName(recorder)
		self.dueBy = dueBy

	def toMessage(self):
		return "*_Week {} - {}_*: {} to record. Share by {}".format(
			self.week, self.topic, self.recorder, self.dueBy)

if __name__ == '__main__':
	main()
