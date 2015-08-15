<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:output version="1.0" encoding="UTF-8"/>

<xsl:key name="QualityGrp" match="Quality" use="text()" />

<xsl:template match="/">
  <xsl:element name="dataroot"><xsl:text>&#xA;</xsl:text>

      <xsl:for-each select="/dataroot/Qualities/Quality[generate-id()
                                = generate-id(key('QualityGrp',text())[1])]">
        <xsl:sort select="translate(key('QualityGrp',text()), 'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
        <Quality type="{key('QualityGrp',text())}" total="{count(key('QualityGrp',text()))}"><xsl:text>&#xA;</xsl:text>
          
           <xsl:for-each select="key('QualityGrp',text())/..">            
            <Character><xsl:value-of select="Character"/></Character><xsl:text>&#xA;</xsl:text>
           </xsl:for-each>
        </Quality><xsl:text>&#xA;</xsl:text>
      </xsl:for-each>
      
  </xsl:element><xsl:text>&#xA;</xsl:text>
</xsl:template>

</xsl:stylesheet>