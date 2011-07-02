import sys
sys.path.append( '..' )

from PyRTF import *

def MergedCells( ) :
	#  another test for the merging of cells in a document
	doc = Document()
	section = Section()
	doc.Sections.append( section )

	#  create the table that will get used for all of the "bordered" content

	col1 = 1000
	col2 = 1000
	col3 = 1000
	col4 = 2000

	section.append( 'Table Two' )

	table = Table( col1, col2, col3 )
	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two'                   ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two', span=2 ) )
	table.AddRow( Cell( 'A-one', span=3 ) )
	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two'                   ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one', span=2 ), Cell( 'A-two' ) )
	section.append( table )

	section.append( 'Table Two' )

	table = Table( col1, col2, col3 )
	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two', vertical_merge=True ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( vertical_merge=True ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two', start_vertical_merge=True ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( vertical_merge=True ), Cell( 'A-three' ) )

	table.AddRow( Cell( Paragraph( ParagraphPS( alignment=ParagraphPS.CENTER ), 'SPREAD' ),
					    span=3 ) )

	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two', vertical_merge=True ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( vertical_merge=True ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( 'A-two', start_vertical_merge=True ), Cell( 'A-three' ) )
	table.AddRow( Cell( 'A-one'   ), Cell( vertical_merge=True ), Cell( 'A-three' ) )

	section.append( table )

	#  
	section.append( 'Table Three' )
	table = Table( col1, col2, col3, col4 )
	table.AddRow( Cell( 'This is pretty amazing', flow=Cell.FLOW_LR_BT, start_vertical_merge=True ),
		          Cell( 'one' ), Cell( 'two' ), Cell( 'three' ) )

	for i in range( 10 ) :	
		table.AddRow( Cell( vertical_merge=True ),
			          Cell( 'one' ), Cell( 'two' ), Cell( 'three' ) )
	section.append( table )

	section.append( 'Table Four' )
	table = Table( col4, col1, col2, col3 )
	table.AddRow( Cell( 'one' ), Cell( 'two' ), Cell( 'three' ),
	              Cell( 'This is pretty amazing', flow=Cell.FLOW_RL_TB, start_vertical_merge=True ) )

	for i in range( 10 ) :	
		table.AddRow( Cell( 'one' ), Cell( 'two' ), Cell( 'three' ),
					  Cell( vertical_merge=True ))
	section.append( table )


	return doc
	
if __name__ == '__main__' :
	renderer = Renderer()
	
	renderer.Write( MergedCells(), file( 'MergedCells.rtf', 'w' ) )
	
	print "Finished"	
	
	
