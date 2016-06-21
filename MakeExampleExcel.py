# coding=utf8

import os
import datetime
import operator
import random
import xlsxwriter
import Utils
from Excel import GetExcelReader
from FitSheetWrapper import FitSheetWrapper

def get_license():
	return u''.join( 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[random.randint(0,25)] for i in xrange(6) )

def get_uci_code():
	dob = datetime.date.today() - datetime.timedelta( days=random.normalvariate(25,3)*365.25 )
	return u'FRA{}'.format( dob.strftime( '%Y%m%d' ) ) 
	
def make_title( s ):
	words = [(w.upper() if w == u'uci' else w) for w in s.split(u'_')]
	return u' '.join( (w[0].upper() + w[1:]) for w in words if w not in (u'of', u'in', u'or') )
	
def MakeExampleExcel():
	random.seed( 0xed )
	
	common_first_names = [unicode(n,'utf-8') for n in 'Léopold Grégoire Aurélien Rémi Léandre Thibault Kylian Nathan Lucas Enzo Léo Louis Hugo Gabriel Ethan Mathis Jules Raphaël Arthur Théo Noah Timeo Matheo Clément Maxime Yanis Maël'.split()]
	common_last_names = [unicode(n,'utf-8') for n in 'Tisserand Lavergne Guignard Parmentier Evrard Leclerc Martin Bernard Dubois Petit Durand Leroy Moreau Simon Laurent Lefevre Roux Fournier Dupont'.split()]
	teams = [unicode(n,'utf-8') for n in 'Pirates of the Pavement,Coastbusters,Tour de Friends,Pesky Peddlers,Spoke & Mirrors'.split(',')]
	
	fname_excel = os.path.join( Utils.getHomeDir(), 'StageRaceGC_Test_Input.xlsx' )
	
	wb = xlsxwriter.Workbook( fname_excel )
	bold_format = wb.add_format( {'bold': True} )
	time_format = wb.add_format( {'num_format': 'hh:mm:ss'} )
	high_precision_time_format = wb.add_format( {'num_format': 'hh:mm:ss.000'} )
	
	ws = wb.add_worksheet('Registration')
	fit_sheet = FitSheetWrapper( ws )
	
	fields = ['bib', 'first_name', 'last_name', 'uci_code', 'license', 'team']
	row = 0
	for c, field in enumerate(fields):
		fit_sheet.write( row, c, make_title(field), bold_format )
		
	riders = 20
	team_size = 4
	bibs = []
	for i in xrange(riders):
		row += 1
		bibs.append((i//team_size+1)*10 + (i%team_size))
		fit_sheet.write( row, 0, bibs[i] )
		fit_sheet.write( row, 1, common_first_names[i%len(common_first_names)] )
		fit_sheet.write( row, 2, common_last_names[i%len(common_last_names)] )
		fit_sheet.write( row, 3, get_uci_code() )
		fit_sheet.write( row, 4, get_license() )
		fit_sheet.write( row, 5, teams[i//team_size] )

	stageCount = 5
	for stage in xrange(stageCount):
		if stage == 3-1:
			tf = high_precision_time_format
			race_time = 60*60
			ws = wb.add_worksheet('Stage {}-ITT'.format(stage+1))
		else:
			tf = time_format
			race_time = 4*60*60
			ws = wb.add_worksheet('Stage {}-RR'.format(stage+1))
		fit_sheet = FitSheetWrapper( ws )
		fields = ['bib', 'time', 'place']
		row = 0
		for c, field in enumerate(fields):
			fit_sheet.write( row, c, make_title(field), bold_format )
		
		bibAB = []
		for i, (bib, t) in enumerate(sorted( ((bib, random.normalvariate(race_time-bib/4.0, 5*60)) for bib in bibs), key=operator.itemgetter(1) )):
			row += 1
			fit_sheet.write( row, 0, bib )
			fit_sheet.write( row, 1, t/(24.0*60.0*60.0), tf )
			if stage in (4-1, 5-1) and i == len(bibs)-1:
				bibAB.append( bib )
				fit_sheet.write( row, 2, 'AB' )
	
		for b in bibAB:
			bibs.remove( b )
	
	wb.close()
	
	return fname_excel
	
if __name__ == '__main__':
	print MakeExampleExcel()
