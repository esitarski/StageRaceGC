import wx
import sys
import datetime
import xlsxwriter
import Utils
import Model
from FitSheetWrapper import FitSheetWrapper

def StageRaceGCToExcel( fname_excel, model ):
	def getRiderInfo( bib ):
		rider = model.registration.bibToRider[bib]
		return u'{}: {}'.format(bib, rider.results_name)
		
	def formatContext( context ):
		lines = []
		for c in context:
			if len(c) == 3:
				t, p, bib = c
				lines.append( u'{}  {} ({})'.format(getRiderInfo(bib), Utils.formatTime(t), Utils.ordinal(p)) )
			elif len(c) == 2:
				p, bib = c
				lines.append( u'{} ({})'.format(getRiderInfo(bib), Utils.ordinal(p)) )
			elif len(c) == 1:
				bib = c[0]
				lines.append( getRiderInfo(bib) )
			else:
				assert False
		return u'\n'.join( lines )

	def formatContextList( context ):
		lines = [formatContext(c).replace('\n', ' - ') for c in context]
		return u'\n'.join( lines )
		
	wb = xlsxwriter.Workbook( fname_excel )
	
	bold_format = wb.add_format( {'bold': True} )
	time_format = wb.add_format( {'num_format': 'hh:mm:ss'} )
	high_precision_time_format = wb.add_format( {'num_format': 'hh:mm:ss.000'} )
	
	comment_style = {'width':400}
	wide_comment_style = {'width':800}
	narrow_comment_style = {'width':256}
	
	#---------------------------------------------------------------------------------------
	def writeIC( ws, stage ):
		fit_sheet = FitSheetWrapper( ws )
		
		ic_fields = Model.IndividualClassification._fields[1:-1]
		riderFields = set( model.registration.getFieldsInUse() )
		headers = (
			['Place', 'Bib', 'Last Name', 'First Name', 'Team'] +
			(['UCI Code'] if 'uci_code' in riderFields else []) +
			(['License'] if 'license' in riderFields else []) +
			[Utils.fieldToHeader(h) for h in ic_fields]
		)
		
		rowNum = 0
		for c, h in enumerate(headers):
			fit_sheet.write( rowNum, c, h, bold_format )
		rowNum +=1
		
		for place, r in enumerate(stage.individual_gc, 1):
			rider = model.registration.bibToRider[r.bib]
		
			col = 0
			if r.retired_stage > 0:
				fit_sheet.write( rowNum, col, 'AB' ); col += 1
			else:
				fit_sheet.write( rowNum, col, place ); col += 1
			
			fit_sheet.write( rowNum, col, r.bib ); col += 1
			fit_sheet.write( rowNum, col, rider.last_name.upper() ); col += 1
			fit_sheet.write( rowNum, col, rider.first_name ); col += 1
			fit_sheet.write( rowNum, col, rider.team ); col += 1
			
			if 'uci_code' in riderFields:
				fit_sheet.write( rowNum, col, rider.team ); col += 1
			if 'license' in riderFields:
				fit_sheet.write( rowNum, col, rider.license ); col += 1
			
			if r.retired_stage == 0:
				fit_sheet.write( rowNum, col, r.total_time_with_bonuses / (24.0*60.0*60.0), time_format ); col += 1
				fit_sheet.write( rowNum, col, r.total_time_with_bonuses_plus_second_fractions / (24.0*60.0*60.0), high_precision_time_format ); col += 1
				fit_sheet.write( rowNum, col, r.last_stage_place ); col += 1
			
			rowNum +=1
	
	#---------------------------------------------------------------------------------------
	def writeTeamClass( ws, stage ):
		fit_sheet = FitSheetWrapper( ws )
		
		headers = ['Place', 'Team', 'Combined Times', 'Combined Places', 'Best Rider GC']
		
		rowNum = 0
		for c, h in enumerate(headers):
			fit_sheet.write( rowNum, c, h, bold_format )
		rowNum +=1
		
		for place, tc in enumerate(stage.team_classification, 1):
			col = 0
			fit_sheet.write( rowNum, col, place ); col += 1
			fit_sheet.write( rowNum, col, tc.team ); col += 1
			
			fit_sheet.write( rowNum, col, tc.sum_best_top_times.value / (24.0*60.0*60.0), time_format )
			# ws.write_comment( rowNum, col, formatContext(tc.sum_best_top_times.context), comment_style )
			col += 1
			
			fit_sheet.write( rowNum, col, tc.sum_best_top_places.value )
			# ws.write_comment( rowNum, col, formatContext(tc.sum_best_top_places.context), comment_style )
			col += 1
			
			fit_sheet.write( rowNum, col, tc.best_place.value )
			# ws.write_comment( rowNum, col, formatContext(tc.best_place.context), comment_style )
			col += 1
			rowNum +=1

	#---------------------------------------------------------------------------------------
	def writeTeamGC( ws ):
		fit_sheet = FitSheetWrapper( ws )
		
		headers = (
			['Place', 'Team', 'Combined Time'] +
			['# {} Places'.format(Utils.ordinal(i+1)) for i in xrange(len(model.all_teams))] +
			['Best Rider GC']
		)
		
		rowNum = 0
		for c, h in enumerate(headers):
			fit_sheet.write( rowNum, c, h, bold_format )
		rowNum +=1
		
		for place, tgc in enumerate(model.team_gc, 1):
			col = 0
			fit_sheet.write( rowNum, col, place ); col += 1
			
			fit_sheet.write( rowNum, col, tgc[-1] ); col += 1
			
			fit_sheet.write( rowNum, col, tgc[0].value / (24.0*60.0*60.0), time_format )
			# ws.write_comment( rowNum, col, formatContextList(tgc[0].context), wide_comment_style )
			col += 1
			
			for i in xrange(1, len(tgc)-2):
				if tgc[i].value:
					fit_sheet.write( rowNum, col, tgc[i].value )
					# ws.write_comment( rowNum, col, u'\n'.join(tgc[i].context), narrow_comment_style )
				col += 1
			
			fit_sheet.write( rowNum, col, tgc[-2].value )
			# ws.write_comment( rowNum, col, formatContext(tgc[-2].context), comment_style )
			col += 1
			
			rowNum +=1
		
		for team in model.unranked_teams:
			col = 0
			fit_sheet.write( rowNum, col, 'DNF' ); col += 1
			fit_sheet.write( rowNum, col, team ); col += 1
			rowNum +=1
	
	#---------------------------------------------------------------------------------------
	if model.stages:
		writeIC(  wb.add_worksheet('IndividualGC'), model.stages[-1] )
	writeTeamGC( wb.add_worksheet('TeamGC') )
	for stage in reversed(model.stages):
		writeIC(  wb.add_worksheet(stage.sheet_name + '-GC'), stage )
		writeTeamClass( wb.add_worksheet(stage.sheet_name + '-TeamClass'), stage )
	
	wb.close()

