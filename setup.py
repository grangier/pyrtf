"""PyRTF - Rich Text Format Document Generation

PyRTF is a pure python module for the efficient generation of rich text format
documents. Supports styles, tables, cell merging, jpg and png images and tries
to maintain compatibility with as many RTF readers as possible. """

classifiers = """\
Development Status :: 4 - Beta
Topic :: Text Editors :: Text Processing
Topic :: Software Development :: Libraries :: Python Modules
Intended Audience :: Developers
Programming Language :: Python
License :: OSI Approved :: GNU General Public License (GPL)
License :: OSI Approved :: GNU Library or Lesser General Public License (LGPL)
"""

import	sys
from distutils.core import setup

if sys.version_info < (2, 3):
    _setup = setup
    def setup(**kwargs):
        if kwargs.has_key("classifiers"):
            del kwargs["classifiers"]
        _setup(**kwargs)

doclines = __doc__.split("\n")

setup( name		= 'PyRTF',
       version	        = '0.45',
       author           = 'Simon Cusack',
       author_email     = 'scusack@sourceforge.net',
       url              = 'http://pyrtf.sourceforge.net/',
       license	        = 'http://www.gnu.org/licenses/gpl.html',
       platforms	= [ 'Any' ],
       description	= doclines[0],
       classifiers	= filter( None, classifiers.split( '\n' ) ),
       long_description = '\n'.join( doclines[2:] ),
       keywords         = ( 'RTF',
                            'Rich Text',
                            'Rich Text Format',
                            'documentation',
                            'reports' ),
       packages         = [ 'PyRTF', ],
       package_dir      = { '' : '.' } )
