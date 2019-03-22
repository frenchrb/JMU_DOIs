<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:datacite="http://datacite.org/schema/kernel-4"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:functx="http://www.functx.com"
    exclude-result-prefixes="xs functx datacite">
    <xsl:output encoding="UTF-8" method="xml" indent="yes"/>
    
    <!-- Transforms Excel Named XML into DataCite XML, JMU Scholarly Commons ETDs (draft DOI)
    Created 2019/03/22 by Rebecca B. French, Metadata Analyst Librarian at James Madison University
    This software is distributed under a Creative Commons Attribution Non-Commercial License -->
    
    <!-- start functx -->
    <!-- Replacements for curly quotes, html tags, en and em dashes -->
    <xsl:variable name="fr" select="('&#8220;', '&#8221;', '&lt;p&gt;', '&lt;/p&gt;', '&lt;em&gt;', '&lt;/em&gt;', '&lt;strong&gt;', '&lt;/strong&gt;', '&#xD;', '&lt;br /&gt;', '&#x2013;', '&#x2014;')"/>
    <xsl:variable name="to" select="('&quot;', '&quot;', '', '', '', '', '', '', '', '', '-', '--')"/>
    
    <!-- Replacements for degrees, titles, etc. -->
    <xsl:variable name="degFr" select="(', Ph.D.', ', PhD', ', Ph. D.', ', Ph.D', 'Dr. ', ', Associate Professor', ', Assistant Professor', 'RD')"/>
    <xsl:variable name="degTo" select="('', '', '', '', '', '', '', '')"/>
    
    <xsl:function name="functx:replace-multi" as="xs:string?"
        xmlns:functx="http://www.functx.com">
        <xsl:param name="arg" as="xs:string?"/>
        <xsl:param name="changeFrom" as="xs:string*"/>
        <xsl:param name="changeTo" as="xs:string*"/>
        
        <xsl:sequence select="
            if (count($changeFrom) > 0)
            then functx:replace-multi(
            replace($arg, $changeFrom[1],
            functx:if-absent($changeTo[1],'')),
            $changeFrom[position() > 1],
            $changeTo[position() > 1])
            else $arg
            "/>
    </xsl:function>
    
    <xsl:function name="functx:if-absent" as="item()*"
        xmlns:functx="http://www.functx.com">
        <xsl:param name="arg" as="item()*"/>
        <xsl:param name="value" as="item()*"/>
        
        <xsl:sequence select="
            if (exists($arg))
            then $arg
            else $value
            "/>
    </xsl:function>
    <!-- end functx -->
    
    <xsl:template match="/">
        <xsl:for-each select="table/row">
            <xsl:if test="doi='' and calc_url!=''">
                <xsl:variable name="setName">
                    <xsl:value-of select="issue"/>
                </xsl:variable>
                <xsl:variable name="pubYear">
                    <xsl:value-of select="substring(publication_date, 1, 4)"/>
                </xsl:variable>
                <xsl:variable name="itemID">
                    <xsl:analyze-string select="calc_url" regex="https://commons\.lib\.jmu\.edu/.*?/">
                        <xsl:non-matching-substring>
                            <xsl:value-of select="."/>
                        </xsl:non-matching-substring>
                    </xsl:analyze-string>
                </xsl:variable>
            
                <xsl:result-document method="xml" href="DataCite_metadata_drafts/etd_{$setName}_{$itemID}_draft.xml">
                    <xsl:call-template name="row"/>
                </xsl:result-document>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
    
    <!--
    <xsl:template match="table">
        <resources>
            <xsl:apply-templates/>
        </resources>
    </xsl:template>
    -->
    
    <xsl:template name="row" match="row">
        <xsl:variable name="givenName">
            <xsl:value-of select="author1_fname"/>
            <xsl:if test="author1_mname != ''">
                <xsl:text> </xsl:text>
                <xsl:value-of select="author1_mname"/>
                <xsl:if test="matches(author1_mname, '[A-Z]$')">
                    <xsl:text>.</xsl:text>
                </xsl:if>
            </xsl:if>
        </xsl:variable>
        
        <xsl:variable name="creatorName">
            <xsl:choose>
                <xsl:when test="preferred_name != ''">
                    <xsl:value-of select="replace(preferred_name, '^(.*) ([''A-z-]+)$', '$2, $1')"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="author1_lname"/>
                    <xsl:text>, </xsl:text>
                    <xsl:value-of select="$givenName"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        
        <xsl:variable name="itemID">
            <xsl:analyze-string select="calc_url" regex="https://commons\.lib\.jmu\.edu/.*?/">
                <xsl:non-matching-substring>
                    <xsl:value-of select="."/>
                </xsl:non-matching-substring>
            </xsl:analyze-string>
        </xsl:variable>
        
        <xsl:variable name="DOI">
            <xsl:text>10.5072/etd/</xsl:text>
            <xsl:value-of select="issue"/>
            <xsl:text>/</xsl:text>
            <!--<xsl:value-of select="substring(publication_date, 1, 4)"/>
                <xsl:text>/</xsl:text>-->
            <xsl:value-of select="$itemID"/>
        </xsl:variable>
        
        <resource xmlns="http://datacite.org/schema/kernel-4"
            xsi:schemaLocation="http://datacite.org/schema/kernel-4 http://schema.datacite.org/meta/kernel-4/metadata.xsd"> 
            <xsl:element name="identifier">
                <xsl:attribute name="identifierType">DOI</xsl:attribute>
                <xsl:value-of select="$DOI"/>
            </xsl:element>
            
            <xsl:element name="creators">
                <xsl:element name="creator">
                    <xsl:element name="creatorName">
                        <xsl:value-of select="$creatorName"/>
                    </xsl:element>
                    <xsl:element name="givenName">
                        <xsl:value-of select="$givenName"/>
                    </xsl:element>
                    <xsl:element name="familyName">
                        <xsl:value-of select="author1_lname"/>
                    </xsl:element>
                    
                    <xsl:if test="orcid != ''">
                        <xsl:element name="nameIdentifier">
                            <xsl:attribute name="schemeURI">https://orcid.org/</xsl:attribute>
                            <xsl:attribute name="nameIdentifierScheme">ORCID</xsl:attribute>
                            <xsl:analyze-string select="orcid" regex="\d{{4}}\-\d{{4}}\-\d{{4}}\-\d{{4}}">
                                <xsl:matching-substring>
                                    <xsl:value-of select="."/>
                                </xsl:matching-substring>
                            </xsl:analyze-string>
                        </xsl:element>
                    </xsl:if>
                    
                    <xsl:element name="affiliation">
                        <xsl:text>James Madison University. </xsl:text>
                        <xsl:value-of select="department"/>
                    </xsl:element>
                </xsl:element>
            </xsl:element>
            
            <xsl:element name="titles">
                <xsl:element name="title">
                    <xsl:value-of select="title"/>
                </xsl:element>
            </xsl:element>
            
            <xsl:element name="publisher">
                <xsl:text>James Madison University Scholarly Commons</xsl:text>
            </xsl:element>
            
            <xsl:element name="publicationYear">
                <xsl:value-of select="substring(publication_date, 1, 4)"/>
            </xsl:element>
            
            <xsl:element name="resourceType">
                <xsl:attribute name="resourceTypeGeneral">Text</xsl:attribute>
                <xsl:text>Dissertation</xsl:text>
            </xsl:element>
            
            <xsl:element name="subjects">
                <xsl:for-each select="tokenize(keywords, ',')">
                    <xsl:element name="subject">
                        <xsl:value-of select="concat(upper-case(substring(normalize-space(.), 1, 1)), substring(normalize-space(.), 2))"/>
                    </xsl:element>
                </xsl:for-each>
                <xsl:for-each select="tokenize(disciplines, ';')">
                    <xsl:element name="subject">
                        <xsl:value-of select="normalize-space(.)"/>
                    </xsl:element>
                </xsl:for-each>
            </xsl:element>
            
            <xsl:element name="contributors">
                <xsl:element name="contributor">
                    <xsl:attribute name="contributorType">Supervisor</xsl:attribute>
                    <xsl:element name="contributorName">
                        <!-- Remove degrees/titles; invert order -->
                        <xsl:choose>
                            <xsl:when test="not(contains(functx:replace-multi(advisor1, $degFr, $degTo), ','))">
                                <xsl:value-of select="replace(functx:replace-multi(advisor1, $degFr, $degTo), '^(.*) ([''A-z-]+)$', '$2, $1')"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="functx:replace-multi(advisor1, $degFr, $degTo)"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:element>
                    <xsl:element name="affiliation">James Madison University</xsl:element>
                </xsl:element>
                <xsl:if test="advisor2 != ''">
                    <xsl:element name="contributor">
                        <xsl:attribute name="contributorType">Supervisor</xsl:attribute>
                        <xsl:element name="contributorName">
                            <!-- Remove degrees/titles; invert order -->
                            <xsl:choose>
                                <xsl:when test="not(contains(functx:replace-multi(advisor2, $degFr, $degTo), ','))">
                                    <xsl:value-of select="replace(functx:replace-multi(advisor2, $degFr, $degTo), '^(.*) ([''A-z-]+)$', '$2, $1')"/>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of select="functx:replace-multi(advisor2, $degFr, $degTo)"/>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:element>
                        <xsl:element name="affiliation">James Madison University</xsl:element>
                    </xsl:element>
                </xsl:if>
                <xsl:if test="advisor3 != ''">
                    <xsl:element name="contributor">
                        <xsl:attribute name="contributorType">Supervisor</xsl:attribute>
                        <xsl:element name="contributorName">
                            <!-- Remove degrees/titles; invert order -->
                            <xsl:choose>
                                <xsl:when test="not(contains(functx:replace-multi(advisor3, $degFr, $degTo), ','))">
                                    <xsl:value-of select="replace(functx:replace-multi(advisor3, $degFr, $degTo), '^(.*) ([''A-z-]+)$', '$2, $1')"/>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of select="functx:replace-multi(advisor3, $degFr, $degTo)"/>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:element>
                        <xsl:element name="affiliation">James Madison University</xsl:element>
                    </xsl:element>
                </xsl:if>
                
                <xsl:element name="contributor">
                    <xsl:attribute name="contributorType">Sponsor</xsl:attribute>
                    <xsl:element name="contributorName">
                        <xsl:text>James Madison University. </xsl:text>
                        <xsl:value-of select="department"/>
                    </xsl:element>
                </xsl:element>
                <xsl:element name="contributor">
                    <xsl:attribute name="contributorType">HostingInstitution</xsl:attribute>
                    <xsl:element name="contributorName">
                        <xsl:text>James Madison University</xsl:text>
                    </xsl:element>
                </xsl:element>
            </xsl:element>
            
            <xsl:element name="dates">
                <xsl:element name="date">
                    <xsl:attribute name="dateType">Copyrighted</xsl:attribute>
                    <xsl:value-of select="substring(publication_date, 1, 10)"/>
                </xsl:element>
                <xsl:element name="date">
                    <xsl:attribute name="dateType">Available</xsl:attribute>
                    <xsl:value-of select="substring(embargo_date, 1, 10)"/>
                </xsl:element>
            </xsl:element>
                
            <xsl:element name="formats">
                <xsl:element name="format">
                    <xsl:value-of select="format"/>
                </xsl:element>
            </xsl:element>
            
            <xsl:element name="rightsList">
                <xsl:element name="rights">
                    <xsl:attribute name="rightsURI">
                        <xsl:value-of select="distribution_license"/>
                    </xsl:attribute>
                    <xsl:if test="distribution_license = 'http://creativecommons.org/licenses/by-nc-nd/4.0/'">
                        <xsl:text>CC BY-NC-ND 4.0</xsl:text>
                    </xsl:if>
                    <!-- TODO text for different licenses? -->
                </xsl:element>
            </xsl:element>
            
            <xsl:element name="descriptions">
                <xsl:element name="description">
                    <xsl:attribute name="descriptionType">Abstract</xsl:attribute>
                    <xsl:value-of select='functx:replace-multi(replace(replace(abstract, "&#8216;", "&apos;"), "&#8217;", "&apos;"),$fr,$to)'/>
                </xsl:element>
                <xsl:element name="description">
                    <xsl:attribute name="descriptionType">Other</xsl:attribute>
                    <xsl:value-of select="degree_name"/>
                    <xsl:text>, </xsl:text>
                    <xsl:value-of select="substring(publication_date, 1, 4)"/>
                </xsl:element>
                <xsl:element name="description">
                    <xsl:attribute name="descriptionType">Other</xsl:attribute>
                    <xsl:text>Recommended Citation: </xsl:text>
                    <xsl:value-of select="$creatorName"/>
                    <xsl:text>, "</xsl:text>
                    <xsl:value-of select="title"/>
                    <xsl:text>" (</xsl:text>
                    <xsl:value-of select="substring(publication_date, 1, 4)"/>
                    <xsl:text>). </xsl:text>
                    <xsl:choose>
                        <xsl:when test="starts-with(issue, 'diss')">
                            <xsl:text>Dissertations</xsl:text>
                        </xsl:when>
                        <xsl:when test="starts-with(issue, 'dnp')">
                            <xsl:text>Doctor of Nursing Practice (DNP) Final Clinical Projects</xsl:text>
                        </xsl:when>
                        <xsl:when test="starts-with(issue, 'edspec')">
                            <xsl:text>Educational Specialist</xsl:text>
                        </xsl:when>
                        <xsl:when test="starts-with(issue, 'master')">
                            <xsl:text>Masters Theses</xsl:text>
                        </xsl:when>
                    </xsl:choose>
                    <xsl:text>. </xsl:text>
                    <xsl:value-of select="$itemID"/>
                    <xsl:text>. https://doi.org/</xsl:text>
                    <xsl:value-of select="$DOI"/>
                </xsl:element>
            </xsl:element>
        </resource>   
            
        
    </xsl:template>
</xsl:stylesheet>
