<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="xml" indent="yes"/>
<xsl:template match="/">
<dataroot>
<xsl:for-each select="//Employees">
<Extensions>
  <LastName>
  <xsl:value-of select="LastName" />
  </LastName>
  <FirstName>
  <xsl:value-of select="FirstName" />
  </FirstName>
  <Extension>
  <xsl:value-of select="Extension" />
  </Extension>
</Extensions>
</xsl:for-each>
</dataroot>
</xsl:template>
</xsl:stylesheet>
