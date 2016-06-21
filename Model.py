
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

class Rider( object ):
	Fields = (
		'bib',
		'first_name', 'last_name',
		'team',
		'uci_code',
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
			
		if self.row:
			try:
				self.row = int(self.row)
			except ValueError:
				self.row = None
				
		if self.last_name:
			self.last_name = unicode(self.last_name).replace(u'*',u'').strip()
			
		if self.first_name:
			self.first_name = unicode(self.first_name).replace(u'*',u'').replace(u'(JR)',u'').strip()
		
		if self.uci_code:
			self.uci_code = unicode(self.uci_code).strip()
			if len(self.uci_code) != 11:
				raise ValueError( u'invalid uci_code: {}'.format(self.uci_code) )
				
		assert self.bib is not None, 'Missing Bib'
				
	@property
	def full_name( self ):
		return u'{} {}'.format( self.first_name, self.last_name )
		
	@property
	def results_name( self ):
		return u','.join( name for name in [self.last_name.upper(), self.first_name] if name )
		
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

class Result( object ):
	Fields = (
		'bib',
		'time',
		'bonus',
		'place',
		'row'
	)
	NumericFields = set([
		'bib', 'row', 'place', 'time',
	])
	
	def __init__( self, **kwargs ):
		for f in self.Fields:
			setattr( self, f, kwargs.get(f, None) )
			
		assert self.bib is not None, "Missing Bib"
		
		self.time = ExcelTimeToSeconds(self.time) or 0.0
		self.integerSeconds = int('{:.3f}'.format(self.time)[:-4])			
		self.bonus = ExcelTimeToSeconds(self.bonus) or 0.0
		
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
			pass
			
	def __repr__( self ):
		return u'Result({})'.format( u','.join( u'{}'.format(getattr(self, a)) for a in self.Fields ) )

reAlpha = re.compile( '[^A-Z]+' )
header_sub = {
	u'RANK':	u'PLACE',
	u'POS':		u'PLACE',
	u'BIBNUM':	u'BIB',
}
def scrub_header( h ):
	h = reAlpha.sub( '', Utils.removeDiacritic(unicode(h)).upper() )
	return header_sub.get(h, h)

def readSheet( reader, sheet_name, header_fields ):
	header_map = {}
	content = []
	errors = []
	for row_number, row in enumerate(reader.iter_list(sheet_name)):
		if not row:
			continue
		
		# Map the column headers to the standard fields.
		if not header_map:
			for c, v in enumerate(row):
				rv = scrub_header( v )
				
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
	
	return content, errors

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
		content, self.errors = readSheet( reader, self.sheet_name, ['name'] + list(Rider.Fields) )
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
			self.errors.append( '{}: Row {}: Missing Bib'.format(self.sheet_name, row['row']) )
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
		content, self.errors = readSheet( reader, self.sheet_name, Result.Fields )
		for c in content:
			self.addRow( c )
		return self.errors
		
	def __len__( self ):
		return len(self.results)
	
class StageITT( Stage ):
	pass
	
class StageTTT( Stage ):
	pass
	
class StageRR( Stage ):
	pass

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
		self.stages = []
		self.reset()
		
	def reset( self ):
		self.team_gc = []
		self.all_teams = set()
		
	def read( self, fname, callbackfunc=None ):
		self.reset()
		self.stages = []
		self.registration = Registration()
		
		reader = GetExcelReader( fname )
		self.registration.read( reader )
		if callbackfunc:
			callbackfunc( self.registration, self.stages )			
		
		for sheet_name in reader.sheet_names():
			if sheet_name.endswith('-ITT'):
				stage = StageITT( sheet_name )
			elif sheet_name.endswith('-TTT'):
				stage = StageTTT( sheet_name )
			elif sheet_name.endswith('-RR'):
				stage = StageRR( sheet_name )
			else:
				continue
			
			errors = stage.read( reader )
			for r in stage.results:
				if r.bib not in self.registration.bibToRider:
					errors.append( '{}: Row {}: Unknown Bib: {}'. format(stage.sheet_name, r.row, r.bib) )
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
				if not isinstance(r.place, int):
					self.retired.add( r.bib )
				if r.bib in self.retired:
					continue
				team = self.registration.bibToRider[r.bib].team
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
			team = self.registration.bibToRider[c.bib].team
			if team not in best_rider_gc:
				best_rider_gc[team] = VC(place, (place, c.bib))
		
		tgc = [ [team_top_times[team]] + team_place_count[team] + [best_rider_gc[team], team] for team in teams ]
		tgc.sort()
		
		self.team_gc = tgc
		self.unranked_teams = sorted( team for team in self.all_teams if team not in teams )
	
	def getGCs( self ):
		self.reset()
		self.getTeamGC()

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
