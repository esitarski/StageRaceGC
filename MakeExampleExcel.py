# coding=utf8

import os
import datetime
import random
import xlsxwriter
import Utils
from Excel import GetExcelReader
from FitSheetWrapper import FitSheetWrapper
from Model import Source, Result
from GetCallups import make_title

def get_license():
	return u''.join( 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[random.randint(0,25)] for i in xrange(6) )

def MakeExampleExcel( include_uci_points=True, include_national_points=True, include_previous_result=True ):
	Year = datetime.date.today().year
	YearAdjust = Year - 2014

	uci_points = Source( os.path.join(Utils.getImageFolder(), 'UCI_Points.xlsx'), 'UCI Points' )
	uci_points.read()
	for r in uci_points.results:
		r.age += YearAdjust
		r.uci_code = u'{}{}'.format(
						r.nation_code,
						(datetime.date(Year-r.age, 1, 1) + datetime.timedelta(days=random.randint(0,364))).strftime('%Y%m%d'),
					)
		r.license = get_license()

	uci_sample = random.sample( uci_points.results, 20 )
	
	common_first_names = [unicode(n,'utf-8') for n in 'Léopold Grégoire Aurélien Rémi Léandre Thibault Kylian Nathan Lucas Enzo Léo Louis Hugo Gabriel Ethan Mathis Jules Raphaël Arthur Théo Noah Timeo Matheo Clément Maxime Yanis Maël'.split()]
	common_last_names = [unicode(n,'utf-8') for n in 'Tisserand Lavergne Guignard Parmentier Evrard Leclerc Martin Bernard Dubois Petit Durand Leroy Moreau Simon Laurent Lefevre Roux Fournier Dupont'.split()]
	
	other_sample = []
	for i in xrange(20):
		other_sample.append( Result(
				first_name=common_first_names[i%len(common_first_names)],
				last_name=common_last_names[i%len(common_last_names)],
				uci_code=u'FRA' + random.choice(uci_points.results).uci_code[3:],
				license=get_license(),
			)
		)
		
	registration = list(uci_sample) + other_sample
	
	bibs = [i for i in xrange(100,200)]
	random.shuffle( bibs )
	
	for i, r in enumerate(registration):
		r.bib = bibs[i]
	
	fname_excel = os.path.join( Utils.getHomeDir(), 'CS_Test_Input.xlsx' )
	
	wb = xlsxwriter.Workbook( fname_excel )
	ws = wb.add_worksheet('Registration')
	fit_sheet = FitSheetWrapper( ws )
	
	fields = ['bib', 'first_name', 'last_name', 'uci_code', 'license']
	for c, field in enumerate(fields):
		fit_sheet.write( 0, c, make_title(field) )
	for r, result in enumerate(registration):
		for c, field in enumerate(fields):
			fit_sheet.write( r+1, c, getattr(result, field) )
	
	if include_uci_points:
		ws = wb.add_worksheet('UCI Points')
		fit_sheet = FitSheetWrapper( ws )
		for c, header in enumerate('Rank	Name	Nation	Age*	Points'.split()):
			fit_sheet.write( 0, c, header )
		for r, result in enumerate(uci_points.results):
			row = r + 1
			fit_sheet.write( row, 0, u'{} ({})'.format(row, row) )
			fit_sheet.write( row, 1, u'{} {}'.format(result.first_name, result.last_name.upper()) )
			fit_sheet.write( row, 2, result.nation )
			fit_sheet.write( row, 3, result.age )
			fit_sheet.write( row, 4, result.points )

	eligible_for_points = other_sample + [rr for rr in uci_points.results if rr.uci_code[:3] == 'FRA']

	if include_national_points:
		ws = wb.add_worksheet('National Points')
		fit_sheet = FitSheetWrapper( ws )
		for c, header in enumerate(['First Name', 'Last Name', 'License', 'Points']):
			fit_sheet.write( 0, c, header )
		
		for r, result in enumerate(random.sample(eligible_for_points, min(len(eligible_for_points),35))):
			row = r + 1
			fit_sheet.write( row, 0, result.first_name )
			fit_sheet.write( row, 1, result.last_name )
			fit_sheet.write( row, 2, result.license )
			fit_sheet.write( row, 3, random.randint(1, 200) )
	
	if include_previous_result:
		ws = wb.add_worksheet( '{} Result'.format(Year-1) )
		fit_sheet = FitSheetWrapper( ws )
		for c, header in enumerate(['Pos', 'First Name', 'Last Name', 'License']):
			fit_sheet.write( 0, c, header )
		for r, result in enumerate(random.sample(eligible_for_points, min(len(eligible_for_points),35))):
			row = r + 1
			fit_sheet.write( row, 0, row )
			fit_sheet.write( row, 1, result.first_name )
			fit_sheet.write( row, 2, result.last_name )
			fit_sheet.write( row, 3, result.license )

	wb.close()
	
	return fname_excel
	
if __name__ == '__main__':
	MakeExampleExcel()
	
	
