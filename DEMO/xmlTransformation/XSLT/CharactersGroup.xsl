<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:output version="1.0" encoding="UTF-8"/>

<xsl:key name="SourceGrp" match="SourceType" use="text()" />

<xsl:template match="/">
  <xsl:element name="dataroot"><xsl:text>&#xA;</xsl:text>

      <xsl:for-each select="/dataroot/Characters/SourceType[generate-id()
                                = generate-id(key('SourceGrp',text())[1])]">
        <xsl:sort select="key('SourceGrp',text())"/>
        <Works sourcetype="{key('SourceGrp',text())}" total="{count(key('SourceGrp',text()))}"><xsl:text>&#xA;</xsl:text>
                            
              <xsl:element name="Characters"><xsl:text>&#xA;</xsl:text>                      
                      <xsl:for-each select="key('SourceGrp',text())/..">
                        <Character>
                            <Name><xsl:value-of select="Character"/></Name><Source><xsl:value-of select="Source"/></Source>
                        </Character><xsl:text>&#xA;</xsl:text>
                      </xsl:for-each>
              </xsl:element><xsl:text>&#xA;</xsl:text>     
              
        </Works><xsl:text>&#xA;</xsl:text>
      </xsl:for-each>
      
  </xsl:element><xsl:text>&#xA;</xsl:text>
</xsl:template>

</xsl:stylesheet>