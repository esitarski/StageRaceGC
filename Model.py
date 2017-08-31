
import os
import re
import sys
import six
import math
import datetime
from collections import defaultdict, namedtuple
from ValueContext import ValueContext as VC
from Excel import GetExcelReader
import Utils

ClimbCategoryLowest = 4

class Rider( object ):
	Fields = (
		'bib',
		'first_name', 'last_name',
		'team',
		'uci_id',
		'license',
		'row'
	)
	NumericFields = set([
		'bib', 'row',
	])

	def __init__( self, **kwargs ):
		if 'name' in kwargs:
			name = kwargs['name'].replace('*','').strip()
			
			# Find the last alpha character.
			cLast = 'C'
			for i in xrange(len(name)-1, -1, -1):
				if name[i].isalpha():
					cLast = name[i]
					break
			
			if cLast == cLast.lower():
				# Assume the name is of the form LAST NAME First Name.
				# Find the last upper-case letter preceeding a space.  Assume that is the last char in the last_name
				j = 0
				i = 0
				while 1:
					i = name.find( u' ', i )
					if i < 0:
						if not j:
							j = len(name)
						break
					cPrev = name[i-1]
					if cPrev.isalpha() and cPrev.isupper():
						j = i
					i += 1
				kwargs['last_name'] = name[:j]
				kwargs['first_name'] = name[j:]
			else:
				# Assume the name field is of the form First Name LAST NAME
				# Find the last lower-case letter preceeding a space.  Assume that is the last char in the first_name
				j = 0
				i = 0
				while 1:
					i = name.find( u' ', i )
					if i < 0:
						break
					cPrev = name[i-1]
					if cPrev.isalpha() and cPrev.islower():
						j = i
					i += 1
				kwargs['first_name'] = name[:j]
				kwargs['last_name'] = name[j:]
		
		for f in self.Fields:
			setattr( self, f, kwargs.get(f, None) )
			
		if self.license is not None:
			self.license = unicode(self.license).strip()
			
		if self.row is not None:
			try:
				self.row = int(self.row)
			except ValueError:
				self.row = None
				
		if self.last_name:
			self.last_name = unicode(self.last_name).replace(u'*',u'').strip()
		self.last_name = self.last_name or u''
			
		if self.first_name:
			self.first_name = unicode(self.first_name).replace(u'*',u'').replace(u'(JR)',u'').strip()
		self.first_name = self.first_name or u''
		
		assert self.bib is not None, 'Missing Bib'
		self.bib = int(self.bib)
				
		if self.uci_id:
			try:
				self.uci_id = int(self.uci_id)
			except:
				pass
			self.uci_id = unicode(self.uci_id).strip()
			if len(self.uci_id) != 11:
				raise ValueError( u'Row {}: invalid uci_id: {} incorrect length'.format(self.row, self.uci_id) )
			if not self.uci_id.isdigit():
				raise ValueError( u'Row {}: invalid uci_id: {} must be all digits'.format(self.row, self.uci_id) )
			if int(self.uci_id[:-2]) % 97 != int(self.uci_id[-2:]):
				raise ValueError( u'Row {}: invalid uci_id: {} checkdigit error'.format(self.row, self.uci_id) )
				
	@property
	def full_name( self ):
		return u'{} {}'.format( self.first_name, self.last_name )
		
	@property
	def results_name( self ):
		return u','.join( name for name in [(self.last_name or u'').upper(), self.first_name] if name )
		
	def __repr__( self ):
		return u'Rider({})'.format(u','.join( u'{}'.format(getattr(self, a)) for a in self.Fields ))

def ExcelTimeToSeconds( t ):
	if t is not None:
		if isinstance(t, six.string_types):
			t = Utils.StrToSeconds( t.replace('"',':').replace("'",':') )
		else:
			# Assume an Excel float number in days.
			t *= 24.0*60.0*60.0
	return t
	
reNonDigit = re.compile( '[^0-9]' )
	
class Result( object ):
	Fields = (
		'bib',
		'time',
		'bonus',
		'penalty',
		'place',
		'row',
		'kom1', 'kom2', 'kom3', 'kom4', 'kom5', 'kom6', 'kom7', 'kom8',
		'sprint1', 'sprint2', 'sprint3', 'sprint4', 'sprint5', 'sprint6', 'sprint7', 'sprint8',
		'stagesprint',
	)
	NumericFields = set([
		'bib', 'row', 'place', 'time', 'bonus', 'penalty',
		'kom1', 'kom2', 'kom3', 'kom4', 'kom5', 'kom6', 'kom7', 'kom8', 'kom9',
		'sprint1', 'sprint2', 'sprint3', 'sprint4', 'sprint5', 'sprint6', 'sprint7', 'sprint8', 'sprint9',
		'stagesprint',
	])
	
	def __init__( self, **kwargs ):
		self.kom = []
		self.sprint = []
		
		try:
			self.stage_sprint = int(kwargs.pop('stagesprint', '0'))
		except:
			self.stage_sprint = 0
		
		for f in self.Fields:
			if f.startswith('kom'):
				break
			setattr( self, f, kwargs.get(f, None) )
		
		def setListValue( a, k, v ):
			try:
				i = int(reNonDigit.sub('', k))
				v = int(v)
			except:
				return
			if i >= 0:
				if len(a) < i:
					a.extend( [0] * (i-len(a)) )
				a[i-1] = v
		
		for k, v in kwargs.iteritems():
			if k.startswith('kom'):
				setListValue( self.kom, k, v )
			elif k.startswith('sprint'):
				setListValue( self.sprint, k, v )
		
		assert self.bib is not None, "Missing Bib"
		
		self.time = ExcelTimeToSeconds(self.time) or 0.0
		self.integerSeconds = int('{:.3f}'.format(self.time)[:-4])			
		self.bonus = ExcelTimeToSeconds(self.bonus) or 0.0
		self.penalty = ExcelTimeToSeconds(self.penalty) or 0.0
		self.time += self.penalty	# Always include the time penalty.
		
		if not self.place:
			self.place = self.row - 1
		else:
			try:
				self.place = int( self.place )
			except:
				pass
		
		if not isinstance( self.place, int ):
			self.place = 'AB'
		
		try:
			self.row = int(self.row)
		except:
			self.row = 0
		
	def __repr__( self ):
		values = [u'{}'.format(getattr(self, a)) for a in self.Fields
			if not a.startswith('sprint') and not a.startswith('kom')]
		kom = 'kom=[{}]'.format(','.join('{}'.format(v) for v in self.kom))
		sprint = 'sprint=[{}]'.format(','.join('{}'.format(v) for v in self.sprint))
		stage_sprint = 'stage_sprint={}'.format( self.stage_sprint )
		return u'Result({}, {}, {}, {})'.format( u','.join( u'{}'.format(getattr(self, a)) for a in self.Fields ),
			kom, sprint, stage_sprint )

reNonAlphaNum = re.compile( '[^A-Z0-9]+' )
header_sub = {
	u'RANK':	u'PLACE',
	u'POS':		u'PLACE',
	u'BIBNUM':	u'BIB',
	u'NUM':		u'BIB',
	u'NUMBER':	u'BIB',
	u'FNAME':	u'FIRST',
	u'LNAME':	u'LAST',
}
def scrub_header( h ):
	h = unicode(h).upper()
	if h.endswith('_NAME') or h.endswith(' NAME'):
		h = h[:-5]
	h = reNonAlphaNum.sub( '', Utils.removeDiacritic(h) )
	return header_sub.get(h, h)

def readSheet( reader, sheet_name, header_fields ):
	header_map = {}
	content = []
	errors = []
	climb_categories = []
	for row_number, row in enumerate(reader.iter_list(sheet_name)):
		if not row:
			continue
		
		# Map the column headers to the standard fields.
		if not header_map:
			for c, v in enumerate(row):
				rv = scrub_header( v )
				if rv.startswith('KOM'):
					if rv.endswith('C'):
						category = rv[-2]
						category = int(category) if category.isdigit() else 0
						rv = rv[:-2]
					else:
						category = ClimbCategoryLowest
					try:
						i = int(reNonDigit.sub('', rv))
						if len(climb_categories) < i:
							climb_categories.extend( [ClimbCategoryLowest] * (i - len(climb_categories)) )
						climb_categories[i-1] = category
					except:
						pass
						
				for h in header_fields:
					hv = scrub_header( h )
					if rv == hv:
						header_map[h] = c
						break						
			continue
	
		# Create a Result from the row.
		row_fields = {}
		for field, column in header_map.iteritems():
			try:
				row_fields[field] = row[column]
			except IndexError:
				pass
		
		row_fields['row'] = row_number + 1
		
		content.append( row_fields )
	
	return content, climb_categories, errors

class Registration( object ):
	def __init__( self, sheet_name = 'Registration' ):
		self.sheet_name = sheet_name
		self.reset()

	def reset( self ):
		self.riders = []
		self.bibToRider = {}
		self.errors = []
	
	def read( self, reader ):
		self.reset()
		content, _, self.errors = readSheet( reader, self.sheet_name, ['name'] + list(Rider.Fields) )
		for row in content:
			try:
				rider = Rider( **row )
			except Exception as e:
				self.errors.append( e )
				continue
			
			self.riders.append( rider )
			self.bibToRider[rider.bib] = rider
		return self.errors
		
	def getFieldsInUse( self ):
		inUse = []
		for f in Rider.Fields:
			if f != 'row':
				for r in self.riders:
					if getattr(r,f,None):
						inUse.append( f )
						break
		return inUse
	
	def empty( self ):
		return not self.riders
	
	def __len__( self ):
		return len(self.riders)

class Stage( object ):
	def __init__( self, sheet_name ):
		self.sheet_name = sheet_name
		self.reset()
		
	def reset( self ):
		self.results = []
		self.errors = []
		
	def addResult( self, result ):
		self.results.append( result )
		if result.place is None:
			result.place = len(self.results)
		
	def addRow( self, row ):
		if 'bib' not in row:
			self.errors.append( 'Row {}: Missing Bib Column'.format(row['row']) )
			return
		try:
			result = Result(**row)
		except Exception as e:
			self.errors.append( e )
			return
		return self.addResult(result)
		
	def empty( self ):
		return not self.results
	
	def read( self, reader ):
		self.reset()
		content, self.climb_categories, self.errors = readSheet( reader, self.sheet_name, Result.Fields )
		for c in content:
			self.addRow( c )
		bad_categories = [c for c in self.climb_categories if not 0 <= c <= 4]
		if bad_categories:
			self.errors.append( 'Unrecognied climb category (must be 4C, 3C, 2C, 1C or HC)' )
		self.climb_categories = [max(min(c, 4), 0) for c in self.climb_categories]
		return self.errors
		
	def __len__( self ):
		return len(self.results)
	
class StageITT( Stage ):
	pass
	
class StageTTT( Stage ):
	pass
	
class StageRR( Stage ):
	pass

class TeamPenalty( object ):
	Fields = (
		'team',
		'penalty',
		'row'
	)
	NumericFields = set([
		'penalty', 'row'
	])
	
	def __init__( self, **kwargs ):
		for f in self.Fields:
			setattr( self, f, kwargs.get(f, None) )
			
		assert self.bib is not None, "Missing Bib"
		
		self.penalty = ExcelTimeToSeconds(self.penalty) or 0.0
		
		try:
			self.row = int(self.row)
		except:
			pass
			
	def __repr__( self ):
		return u'TeamPenalty({})'.format( u','.join( u'{}'.format(getattr(self, a)) for a in self.Fields ) )

class TeamPenalties( object ):
	def __init__( self, sheet_name = 'Team Penalties' ):
		self.sheet_name = sheet_name
		self.reset()
		
	def reset( self ):
		self.team_penalties = defaultdict(float)
		self.errors = []
		
	def addTeamPenalty( self, team_penalties ):
		self.team_penalties[team_penalties.team] += team_penalties.penalty
		
	def addRow( self, row ):
		if 'team' not in row:
			self.errors.append( 'Row {}: Missing Team Column'.format(row['row']) )
			return
		try:
			team_penalties = TeamPenalty(**row)
		except Exception as e:
			self.errors.append( e )
			return
		return self.addTeamPenalty(team_penalties)
		
	def empty( self ):
		return not self.team_penalties
	
	def read( self, reader ):
		self.reset()
		content, self.errors = readSheet( reader, self.sheet_name, TeamPenalty.Fields )
		for c in content:
			self.addRow( c )
		return self.errors
		
	def __len__( self ):
		return len(self.team_penalties)

IndividualClassification = namedtuple( 'IndividualClassification', [
		'retired_stage',
		'total_time_with_bonuses',
		'total_time_with_bonuses_plus_second_fractions',
		'last_stage_place',
		'bib',
	]
)

TeamClassification = namedtuple( 'TeamClassification', [
		'sum_best_top_times',
		'sum_best_top_places',
		'best_place',
		'team',
	]
)

class Model( object):
	def __init__( self ):
		self.registration = Registration()
		self.team_penalties = TeamPenalties()
		self.stages = []
		self.reset()
		
	def reset( self ):
		self.team_gc = []
		self.sprint_gc = []
		self.kom_gc = []
		self.all_teams = set()
		
	def read( self, fname, callbackfunc=None ):
		self.reset()
		self.stages = []
		self.registration = Registration()
		self.team_penalties = TeamPenalties()
		
		reader = GetExcelReader( fname )
		self.registration.read( reader )
		if callbackfunc:
			callbackfunc( self.registration, self.stages )			
		
		for sheet_name in reader.sheet_names():
			if sheet_name.endswith('-RR'):
				stage = StageRR( sheet_name )
			elif sheet_name.endswith('-ITT'):
				stage = StageITT( sheet_name )
			elif sheet_name.endswith('-TTT'):
				stage = StageTTT( sheet_name )
			elif sheet_name.lower().replace(' ', '') == 'teampenalties':
				self.team_penalties = TeamPenalties( sheet_name )
				errors = team_penalties.read()
				self.all_teams = { r.team for r in self.registration.riders }
				for team in team_penalties.team_penalties.iterkeys():
					if team not in self.all_teams:
						errors.append( 'Unknown Team: {}'. format(team) )
				continue
			else:
				continue
			
			errors = stage.read( reader )
			for r in stage.results:
				if r.bib not in self.registration.bibToRider:
					errors.append( 'Row {}: Unknown Bib: {}'. format(r.row, r.bib) )
			self.stages.append( stage )
			
			if callbackfunc:
				callbackfunc( self.registration, self.stages )			

	def getIndividualGC( self, stageLast = None ):
		self.all_teams = { r.team for r in self.registration.riders }
		
		stageLast = stageLast or self.stages[-1]
		
		# Get all retired riders.
		stageLast.retired = set()
		ic = []
		for i, stage in enumerate(self.stages, 1):
			for r in stage.results:
				if not isinstance(r.place, int) and r not in stageLast.retired:
					stageLast.retired.add( r.bib )
					ic.append( IndividualClassification(i, 0, 0, 0, r.bib) )
			if stage == stageLast:
				break

		# Calculate the classification criteria.
		stageLast.bibs = set()
		stageLast.total_time_without_bonuses = defaultdict( float )
		stageLast.total_time_with_bonuses = defaultdict( float )
		stageLast.total_time_with_bonuses_plus_second_fractions = defaultdict( float )
		stageLast.last_stage_place = defaultdict( int )
		for stage in self.stages:
			for r in stage.results:
				if r.bib in stageLast.retired:
					continue
				stageLast.bibs.add( r.bib )
				
				time_without_bonus = r.integerSeconds
				time_with_bonus = time_without_bonus - r.bonus
				time_with_bonuses_plus_second_fractions = r.time - r.bonus
				
				stageLast.total_time_without_bonuses[r.bib] += time_without_bonus
				stageLast.total_time_with_bonuses[r.bib] += time_with_bonus
				stageLast.total_time_with_bonuses_plus_second_fractions[r.bib] += \
					time_with_bonuses_plus_second_fractions if isinstance(stage, (StageITT, StageTTT)) else time_with_bonus
				
				if stage == stageLast:
					stageLast.last_stage_place[r.bib] = r.place
			
			if stage == stageLast:
				break

		# Populate the result.
		for bib in stageLast.bibs:
			ic.append( IndividualClassification(
					0,
					stageLast.total_time_with_bonuses[bib],
					stageLast.total_time_with_bonuses_plus_second_fractions[bib],
					stageLast.last_stage_place[bib],
					bib
				)
			)

		# Sort to get the unique classification.
		ic.sort( key = lambda c: (
				c.retired_stage,
				c.total_time_with_bonuses,
				c.total_time_with_bonuses_plus_second_fractions,
				c.last_stage_place,
				c.bib,
			)
		)
		stageLast.individual_gc = ic
		
	def getTeamStageClassifications( self ):
		self.retired = set()
		
		for stage in self.stages:
			self.getIndividualGC( stage )
		
			sum_best_top_times = {team: VC() for team in self.all_teams}
			sum_best_top_places = {team: VC() for team in self.all_teams}
			best_place = {}
			top_count = {team: 0 for team in self.all_teams}
			
			for r in stage.results:
				rider = self.registration.bibToRider.get(r.bib, None)
				if not rider:
					continue
				if not isinstance(r.place, int):
					self.retired.add( r.bib )
				if r.bib in self.retired:
					continue
				
				team = rider.team
				if top_count[team] == 3:
					continue
					
				if top_count[team] == 0:
					best_place[team] = VC(r.place, (r.place, r.bib))
				sum_best_top_times[team] += VC(r.integerSeconds, (r.integerSeconds, r.place, r.bib))
				sum_best_top_places[team] += VC(r.place, (r.place, r.bib))
				top_count[team] += 1
			
			stage.team_classification = [
				TeamClassification(sum_best_top_times[team], sum_best_top_places[team], best_place[team], team,)
					for team in sum_best_top_times.iterkeys() if top_count[team] == 3
			]
			
			stage.team_classification.sort()
		
	def getTeamGC( self ):
		self.team_gc = []
		self.unranked_teams = []
		
		if not self.stages:
			return
		
		self.getTeamStageClassifications()
		
		total_teams = len( self.all_teams )
		
		teams = set()
		for stage in reversed(self.stages):
			try:
				teams = { tc.team for tc in stage.team_classification }
				break
			except AttributeError:
				continue
			
		team_top_times = { team: VC() for team in teams }
		team_place_count = { team:  [VC() for i in xrange(total_teams)] for team in teams }
		
		for stage in self.stages:
			try:
				team_classification = stage.team_classification
			except AttributeError:
				continue
				
			for place, tc in enumerate(team_classification, 1):
				if tc.team in team_top_times:
					team_top_times[tc.team] += VC(tc.sum_best_top_times.value, [tc.sum_best_top_times.context])
					team_place_count[tc.team][place-1] += VC(1, stage.sheet_name)
		
		best_rider_gc = {}
		for place, c in enumerate(self.stages[-1].individual_gc, 1):
			rider = self.registration.bibToRider.get( c.bib, None )
			if not rider:
				continue
			team = rider.team
			if team not in best_rider_gc:
				best_rider_gc[team] = VC(place, (place, c.bib))
		
		tgc = [ [team_top_times[team] + VC(self.team_penalties.team_penalties[team])] +
					team_place_count[team] + [best_rider_gc[team], team]
			for team in teams ]
		tgc.sort()
		
		self.team_gc = tgc
		self.unranked_teams = sorted( team for team in self.all_teams if team not in teams )
	
	def getKOMGC( self ):
		self.kom_gc = []
		if not self.stages:
			return
		
		lastStage = self.stages[-1]
		retired = lastStage.retired
		lastStageGC = {ic.bib: place for place, ic in enumerate(lastStage.individual_gc, 1) if ic.retired_stage == 0}
		
		bibCategoryWins = defaultdict(lambda: [0]*(ClimbCategoryLowest+1))
		bibKOMTotal = defaultdict(int)
		for stage in self.stages:
			if not stage.climb_categories:
				continue
			climb_winner = [0] * len(stage.climb_categories)
			climb_points = [0] * len(stage.climb_categories)
			for result in stage.results:
				if result.bib in retired:
					continue
				for i, v in enumerate(result.kom):
					bibKOMTotal[result.bib] += v
					if v > climb_points[i]:
						climb_winner[i] = result.bib
						climb_points[i] = v
			for bib, category in zip(climb_winner, stage.climb_categories):
				if bib:
					bibCategoryWins[bib][category] += 1
		# Sort by decreasing total KOM, then by decreasing wins by categegory climb, then by gc.
		kom = [[bibKOMTotal[bib]] + bibCategoryWins[bib] + [lastStageGC[bib], bib]
			for bib in bibKOMTotal.iterkeys()]
		kom.sort( reverse=True, key=lambda x: x[:-2] + [-x[-2]] )
		self.kom_gc = kom
		
	def getSprintGC( self ):
		self.sprint_gc = []
		if not self.stages:
			return
		
		lastStage = self.stages[-1]
		retired = lastStage.retired
		lastStageGC = {ic.bib: place for place, ic in enumerate(lastStage.individual_gc, 1) if ic.retired_stage == 0}
		
		bibStageWins = defaultdict(int)
		bibSprintWins = defaultdict(int)
		bibSprintTotal = defaultdict(int)
		for stage in self.stages:
			if not isinstance(stage, StageRR):
				continue
			try:
				stage_sprint_count = max(len(r.sprint) for r in stage.results)
			except:
				continue
			stage_sprint_points = 0
			stage_sprint_winner = 0
			sprint_points = [0] * stage_sprint_count
			sprint_winner = [0] * stage_sprint_count
			for result in stage.results:
				if result.bib in retired:
					continue
				
				bibSprintTotal[result.bib] += result.stage_sprint
				if result.stage_sprint > stage_sprint_points:
					stage_sprint_points = result.stage_sprint
					stage_sprint_winner = result.bib
					
				for i, v in enumerate(result.sprint):
					bibSprintTotal[result.bib] += v
					if v > sprint_points[i]:
						sprint_winner[i] = result.bib
						sprint_points[i] = v
			for bib in sprint_winner:
				if bib:
					bibSprintWins[bib] += 1
			if stage_sprint_winner:
				bibStageWins[stage_sprint_winner] += 1
		
		sprint = [[bibSprintTotal[bib], bibStageWins[bib], bibSprintWins[bib], lastStageGC[bib], bib]
			for bib in bibSprintTotal.iterkeys()]
		
		# Sort by decreasing total Sprint points, then by stage wins, then by sprint wins, then by gc.
		sprint.sort( reverse=True, key=lambda x: x[:-2] + [-x[-2]] )
		self.sprint_gc = sprint
	
	def getGCs( self ):
		self.reset()
		self.getTeamGC()
		self.getKOMGC()
		self.getSprintGC()

model = None
def read( fname, callbackfunc=None ):
	global model
	model = Model()
	model.read( fname, callbackfunc=callbackfunc )
	return model

def unitTest():
	fname = 'StageRaceGCTest.xlsx' 
	model = Model()
	model.read( fname )
	#print 'Registration:', len(model.registration.riders)
	#print model.registration.riders
	
	model.getGCs()
	
	print '*' * 70
	print 'Individual GC'
	print '*' * 70
	for gc in model.stages[-1].individual_gc:
		print gc
		
	print '*' * 70
	print 'Team GC'
	print '*' * 70
	for gc in model.team_gc:
		print gc
		
	print '*' * 70
	print 'Team Classification by Stage'
	print '*' * 70
	for stage in model.stages:
		for gc in stage.team_classification:
			print gc
		print '-----------------'
	return model, fname
	
if __name__ == '__main__':
	unitTest()
