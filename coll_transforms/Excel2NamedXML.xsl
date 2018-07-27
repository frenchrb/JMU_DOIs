<?xml version='1.0'?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" exclude-result-prefixes="ss">

	<!--
	Transforms a tabular list from a single Excel worksheet from Excel's native XML format to a more useful one with named nodes (essentially the XML equivalent of a CSV file).
	Node names are taken from the corresponding cell name or content of the first row (assumed to be a "header" row), or the cell name of the current cell, in that order.
	White space is replaced with a single "_" to attempt valid node names.
	Parameters can be passed to choose a different than the first worksheet within a workbook, the number of header rows to skip (if any), and to override the skipping of empty rows.

	Change History
	==============
	08JUN14 - extend the list of characters not permitted (and thus translated) for a node name.
	10JUN14 - option to skip empty rows (most useful behind last data row)

	© 2014, Andy Schmidt, ASchmidt@Anamera.net 
	-->

	<xsl:param name="nWorksheet" select="1"/>
	<!-- Worksheet # within a workbook -->
	<xsl:param name="nSkipRows" select="1"/>
	<!-- # of header row(s) to skip -->
	<xsl:param name="nSkipEmpty" select="1"/>
	<!-- 1 = Skip rows with zero string lengths -->

	<xsl:output method="xml" version="1.0" encoding="utf-8" media-type="application/xml"
		indent="yes"/>
	<xsl:strip-space elements="*"/>

	<xsl:template match="/">
		<xsl:element name="table">
			<xsl:comment>Transformed by Excel2NamedXML, © 2014, Andy Schmidt, ASchmidt@Anamera.net</xsl:comment>
			<xsl:apply-templates select="//parent::node()[$nWorksheet]/ss:Row"/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="ss:Row">
		<xsl:if
			test="position() &gt; $nSkipRows and (nSkipEmpty != '1' or count(ss:Cell/ss:Data/text()[string-length()>0]) > 0)">
			<xsl:element name="row">
				<xsl:attribute name="number">
					<!-- Show row number for easy reference -->
					<xsl:value-of select="position()-$nSkipRows"/>
				</xsl:attribute>
				<xsl:apply-templates select="ss:Cell"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>

	<xsl:template match="ss:Cell">
		<xsl:variable name="nPosition" select="position()"/>
		<xsl:element
			name="{translate(normalize-space((//parent::node()[$nWorksheet]/ss:Row[1]/ss:Cell[$nPosition]/ss:NamedCell/@ss:Name | //parent::node()[$nWorksheet]/ss:Row[1]/ss:Cell[$nPosition]/ss:Data[not(parent::node()/ss:NamedCell/@ss:Name)] | ss:NamedCell/@ss:Name)[1]),' #;~`!@$%^*+=|:,?()[]&lt;&gt;&amp;&quot;&#10;&#13;','_-.--.--------...--------..')}">
			<xsl:if test="string-length(ss:NamedCell/@ss:Name[1]) &gt; 0">
				<xsl:attribute name="name">
					<!-- Show name for this individual cell -->
					<xsl:value-of select="ss:NamedCell/@ss:Name[1]"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:value-of select="ss:Data"/>
		</xsl:element>
	</xsl:template>

</xsl:stylesheet>
