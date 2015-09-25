#!/usr/bin/env python

from AppKit import NSURL
from ScriptingBridge import SBApplication
import sys

DEBUG = False
BUNDLE = 'com.apple.iWork.Keynote'

SAVING_OPTIONS = {
    'yes': 0x79657320,  # 'yes '
    'no':  0x6E6F2020,  # 'no  '
    'ask': 0x61736B20,  # 'ask '
}

EXPORT_FORMAT = {
    'Khtm': 0x4B68746D, # HTML
    'Kmov': 0x4B6D6F76, # QuickTime movie
    'Kpdf': 0x4B706466, # PDF
    'Kimg': 0x4B696D67, # Image
    'Kppt': 0x4B707074, # Microsoft PowerPoint
    'Kkey': 0x4B6B6579, # Keynote 09
}

if len(sys.argv) < 2:
    print "usage: %s <keynote-file>" % sys.argv[0]
    sys.exit(-1)

# Export options
keynote_file = sys.argv[1]
to_file = NSURL.fileURLWithPath_(keynote_file.split('.key')[0] + '.pdf')
as_format = EXPORT_FORMAT['Kpdf']
with_properties = None

if DEBUG:
    print("   KEYNOTE_FILE: %s" % keynote_file)
    print("        TO_FILE: %s" % to_file)
    print("      AS_FORMAT: %s" % as_format)
    print("WITH_PROPERTIES: %s" % with_properties)

# Open Keynote file
keynote = SBApplication.applicationWithBundleIdentifier_(BUNDLE)
doc = keynote.open_(keynote_file)

# Export to format
doc.exportTo_as_withProperties_(to_file, as_format, with_properties)

# Close keynote
doc.closeSaving_savingIn_(SAVING_OPTIONS['no'], None)
keynote.quitSaving_(SAVING_OPTIONS['no'])
