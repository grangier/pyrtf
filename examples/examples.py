#!/usr/bin/env python

import sys
sys.path.append( '../' )

from PyRTF import *

SAMPLE_PARA = """The play opens one year after the death of Richard II, and
King Henry is making plans for a crusade to the Holy Land to cleanse himself
of the guilt he feels over the usurpation of Richard's crown. But the crusade
must be postponed when Henry learns that Welsh rebels, led by Owen Glendower,
have defeated and captured Mortimer. Although the brave Henry Percy, nicknamed
Hotspur, has quashed much of the uprising, there is still much trouble in
Scotland. King Henry has a deep admiration for Hotspur and he longs for his
own son, Prince Hal, to display some of Hotspur's noble qualities. Hal is more
comfortable in a tavern than on the battlefield, and he spends his days
carousing with riff-raff in London. But King Henry also has his problems with
the headstrong Hotspur, who refuses to turn over his prisoners to the state as
he has been so ordered. Westmoreland tells King Henry that Hotspur has many of
the traits of his uncle, Thomas Percy, the Earl of Worcester, and defying
authority runs in the family."""

def MakeExample1() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	#	text can be added directly to the section
	#	a paragraph object is create as needed
	section.append( 'Example 1' )

	#	blank paragraphs are just empty strings
	section.append( '' )

	#	a lot of useful documents can be created
	#	with little more than this
	section.append( 'A lot of useful documents can be created '
					'in this way, more advance formating is available '
					'but a lot of users just want to see their data come out '
					'in something other than a text file.' )
	return doc

def MakeExample2() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	#	things really only get interesting after you
	#	start to use styles to spice things up
	p = Paragraph( ss.ParagraphStyles.Heading1 )
	p.append( 'Example 2' )
	section.append( p )

	p = Paragraph( ss.ParagraphStyles.Normal )
	p.append( 'In this case we have used a two styles. '
			  'The first paragraph is marked with the Heading1 style so that it '
			  'will appear differently to the user. ')
	section.append( p )

	p = Paragraph()
	p.append( 'Notice that after I have changed the style of the paragraph '
			  'all subsequent paragraphs have that style automatically. '
			  'This saves typing and is the default behaviour for RTF documents.' )
	section.append( p )

	p = Paragraph()
	p.append( 'I also happen to like Arial so our base style is Arial not Times New Roman.' )
	section.append( p )

	p = Paragraph()
	p.append( 'It is also possible to provide overrides for element of a style. ',
			  'For example I can change just the font ',
			  TEXT( 'size', size=48 ),
			  ' or ',
			  TEXT( 'typeface', font=ss.Fonts.Impact ) ,
			  '.' )
	section.append( p )

	p = Paragraph()
	p.append( 'The paragraph itself can also be overridden in lots of ways, tabs, '
			  'borders, alignment, etc can all be modified either in the style or as an '
			  'override during the creation of the paragraph. '
			  'The next paragraph demonstrates custom tab widths and embedded '
			  'carriage returns, ie new line markers that do not cause a paragraph break.' )
	section.append( p )

	#	ParagraphPS is an alias for ParagraphPropertySet
	para_props = ParagraphPS( tabs = [ TabPS( width=TabPS.DEFAULT_WIDTH     ),
									   TabPS( width=TabPS.DEFAULT_WIDTH * 2 ),
									   TabPS( width=TabPS.DEFAULT_WIDTH     ) ] )
	p = Paragraph( ss.ParagraphStyles.Normal, para_props )
	p.append( 'Left Word', TAB, 'Middle Word', TAB, 'Right Word', LINE,
			  'Left Word', TAB, 'Middle Word', TAB, 'Right Word' )
	section.append( p )

	section.append( 'The alignment of tabs and style can also be controlled. '
					'The following paragraph demonstrates how to use flush right tabs'
					'and leader dots.' )

	para_props = ParagraphPS( tabs = [ TabPS( section.TwipsToRightMargin(),
											  alignment = TabPS.RIGHT,
											  leader    = TabPS.DOTS  ) ] )
	p = Paragraph( ss.ParagraphStyles.Normal, para_props )
	p.append( 'Before Dots', TAB, 'After Dots' )
	section.append( p )

	section.append( 'Paragraphs can also be indented, the following is all at the '
					'same indent level and the one after it has the first line '
					'at a different indent to the rest.  The third has the '
					'first line going in the other direction and is also separated '
					'by a page break.  Note that the '
					'FirstLineIndent is defined as being the difference from the LeftIndent.' )

	section.append( 'The following text was copied from http://www.shakespeare-online.com/plots/1kh4ps.html.' )

	para_props = ParagraphPS()
	para_props.SetLeftIndent( TabPropertySet.DEFAULT_WIDTH *  3 )
	p = Paragraph( ss.ParagraphStyles.Normal, para_props )
	p.append( SAMPLE_PARA )
	section.append( p )

	para_props = ParagraphPS()
	para_props.SetFirstLineIndent( TabPropertySet.DEFAULT_WIDTH * -2 )
	para_props.SetLeftIndent( TabPropertySet.DEFAULT_WIDTH *  3 )
	p = Paragraph( ss.ParagraphStyles.Normal, para_props )
	p.append( SAMPLE_PARA )
	section.append( p )

	#	do a page
	para_props = ParagraphPS()
	para_props.SetPageBreakBefore( True )
	para_props.SetFirstLineIndent( TabPropertySet.DEFAULT_WIDTH )
	para_props.SetLeftIndent( TabPropertySet.DEFAULT_WIDTH )
	p = Paragraph( ss.ParagraphStyles.Normal, para_props )
	p.append( SAMPLE_PARA )
	section.append( p )

	return doc


def MakeExample3() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	p = Paragraph( ss.ParagraphStyles.Heading1 )
	p.append( 'Example 3' )
	section.append( p )

	# changes what is now the default style of Heading1 back to Normal
	p = Paragraph( ss.ParagraphStyles.Normal )
	p.append( 'Example 3 demonstrates tables, tables represent one of the '
			  'harder things to control in RTF as they offer alot of '
			  'flexibility in formatting and layout.' )
	section.append( p )

	section.append( 'Table columns are specified in widths, the following example '
					'consists of a table with 3 columns, the first column is '
					'7 tab widths wide, the next two are 3 tab widths wide. '
					'The widths chosen are arbitrary, they do not have to be '
					'multiples of tab widths.' )

	table = Table( TabPS.DEFAULT_WIDTH * 7,
				   TabPS.DEFAULT_WIDTH * 3,
				   TabPS.DEFAULT_WIDTH * 3 )
	c1 = Cell( Paragraph( 'Row One, Cell One'   ) )
	c2 = Cell( Paragraph( 'Row One, Cell Two'   ) )
	c3 = Cell( Paragraph( 'Row One, Cell Three' ) )
	table.AddRow( c1, c2, c3 )

	c1 = Cell( Paragraph( ss.ParagraphStyles.Heading2, 'Heading2 Style'   ) )
	c2 = Cell( Paragraph( ss.ParagraphStyles.Normal, 'Back to Normal Style'   ) )
	c3 = Cell( Paragraph( 'More Normal Style' ) )
	table.AddRow( c1, c2, c3 )

	c1 = Cell( Paragraph( ss.ParagraphStyles.Heading2, 'Heading2 Style'   ) )
	c2 = Cell( Paragraph( ss.ParagraphStyles.Normal, 'Back to Normal Style'   ) )
	c3 = Cell( Paragraph( 'More Normal Style' ) )
	table.AddRow( c1, c2, c3 )

	section.append( table )
	section.append( 'Different frames can also be specified for each cell in the table '
					'and each frame can have a different width and style for each border.' )

	thin_edge  = BorderPS( width=20, style=BorderPS.SINGLE )
	thick_edge = BorderPS( width=80, style=BorderPS.SINGLE )

	thin_frame  = FramePS( thin_edge,  thin_edge,  thin_edge,  thin_edge )
	thick_frame = FramePS( thick_edge, thick_edge, thick_edge, thick_edge )
	mixed_frame = FramePS( thin_edge,  thick_edge, thin_edge,  thick_edge )

	table = Table( TabPS.DEFAULT_WIDTH * 3, TabPS.DEFAULT_WIDTH * 3, TabPS.DEFAULT_WIDTH * 3 )
	c1 = Cell( Paragraph( 'R1C1' ), thin_frame )
	c2 = Cell( Paragraph( 'R1C2' ) )
	c3 = Cell( Paragraph( 'R1C3' ), thick_frame )
	table.AddRow( c1, c2, c3 )

	c1 = Cell( Paragraph( 'R2C1' ) )
	c2 = Cell( Paragraph( 'R2C2' ) )
	c3 = Cell( Paragraph( 'R2C3' ) )
	table.AddRow( c1, c2, c3 )

	c1 = Cell( Paragraph( 'R3C1' ), mixed_frame )
	c2 = Cell( Paragraph( 'R3C2' ) )
	c3 = Cell( Paragraph( 'R3C3' ), mixed_frame )
	table.AddRow( c1, c2, c3 )

	section.append( table )

	section.append( 'In fact frames can be applied to paragraphs too, not just cells.' )

	p = Paragraph( ss.ParagraphStyles.Normal, thin_frame )
	p.append( 'This whole paragraph is in a frame.' )
	section.append( p )
	return doc

def MakeExample4() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	section.append( 'Example 4' )
	section.append( 'This example test changing the colour of fonts.' )

	#
	#	text properties can be specified in two ways, either a
	#	Text object can have its text properties specified like:
	tps = TextPS( colour=ss.Colours.Red )
	text = Text( 'RED', tps )
	p = Paragraph()
	p.append( 'This next word should be in ', text )
	section.append( p )

	#	or the shortcut TEXT function can be used like:
	p = Paragraph()
	p.append( 'This next word should be in ', TEXT( 'Green', colour=ss.Colours.Green ) )
	section.append( p )

	#	when specifying colours it is important to use the colours from the
	#	style sheet supplied with the document and not the StandardColours object
	#	each document get its own copy of the stylesheet so that changes can be
	#	made on a document by document basis without mucking up other documents
	#	that might be based on the same basic stylesheet

	return doc


#
#  DO SOME WITH HEADERS AND FOOTERS
def MakeExample5() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	section.Header.append( 'This is the header' )
	section.Footer.append( 'This is the footer' )

	p = Paragraph( ss.ParagraphStyles.Heading1 )
	p.append( 'Example 5' )
	section.append( p )

	#	blank paragraphs are just empty strings
	section.append( '' )

	p = Paragraph( ss.ParagraphStyles.Normal )
	p.append( 'This document has a header and a footer.' )
	section.append( p )

	return doc

def MakeExample6() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	section.FirstHeader.append( 'This is the header for the first page.' )
	section.FirstFooter.append( 'This is the footer for the first page.' )

	section.Header.append( 'This is the header that will appear on subsequent pages.' )
	section.Footer.append( 'This is the footer that will appear on subsequent pages.' )

	p = Paragraph( ss.ParagraphStyles.Heading1 )
	p.append( 'Example 6' )
	section.append( p )

	#	blank paragraphs are just empty strings
	section.append( '' )

	p = Paragraph( ss.ParagraphStyles.Normal )
	p.append( 'This document has different headers and footers for the first and then subsequent pages. '
			  'If you insert a page break you should see a different header and footer.' )
	section.append( p )

	return doc

def MakeExample7() :
	doc     = Document()
	ss      = doc.StyleSheet
	section = Section()
	doc.Sections.append( section )

	section.FirstHeader.append( 'This is the header for the first page.' )
	section.FirstFooter.append( 'This is the footer for the first page.' )

	section.Header.append( 'This is the header that will appear on subsequent pages.' )
	section.Footer.append( 'This is the footer that will appear on subsequent pages.' )

	p = Paragraph( ss.ParagraphStyles.Heading1 )
	p.append( 'Example 7' )
	section.append( p )

	p = Paragraph( ss.ParagraphStyles.Normal )
	p.append( 'This document has different headers and footers for the first and then subsequent pages. '
			  'If you insert a page break you should see a different header and footer.' )
	section.append( p )

	p = Paragraph( ss.ParagraphStyles.Normal, ParagraphPS().SetPageBreakBefore( True ) )
	p.append( 'This should be page 2 '
			  'with the subsequent headers and footers.' )
	section.append( p )

	section = Section( break_type=Section.PAGE, first_page_number=1 )
	doc.Sections.append( section )

	section.FirstHeader.append( 'This is the header for the first page of the second section.' )
	section.FirstFooter.append( 'This is the footer for the first page of the second section.' )

	section.Header.append( 'This is the header that will appear on subsequent pages for the second section.' )
	p = Paragraph( 'This is the footer that will appear on subsequent pages for the second section.', LINE )
	p.append( PAGE_NUMBER, ' of ', SECTION_PAGES )
	section.Footer.append( p )

	section.append( 'This is the first page' )

	p = Paragraph( ParagraphPS().SetPageBreakBefore( True ), 'This is the second page' )
	section.append( p )

	return doc

def OpenFile( name ) :
	return file( '%s.rtf' % name, 'w' )

if __name__ == '__main__' :
	DR = Renderer()

	doc1 = MakeExample1()
	doc2 = MakeExample2()
	doc3 = MakeExample3()
	doc4 = MakeExample4()
	doc5 = MakeExample5()
	doc6 = MakeExample6()
	doc7 = MakeExample7()

	DR.Write( doc1, OpenFile( '1' ) )
	DR.Write( doc2, OpenFile( '2' ) )
	DR.Write( doc3, OpenFile( '3' ) )
	DR.Write( doc4, OpenFile( '4' ) )
	DR.Write( doc5, OpenFile( '5' ) )
	DR.Write( doc6, OpenFile( '6' ) )
	DR.Write( doc7, OpenFile( '7' ) )

	print "Finished"
