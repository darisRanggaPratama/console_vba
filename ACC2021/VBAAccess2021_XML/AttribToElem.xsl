<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:rs="urn:schemas-microsoft-com:rowset">
<xsl:output method="xml" encoding="UTF-8" />

  <xsl:template match="/">

  <!-- root element for the XML output -->
  <Products xmlns:z="#RowsetSchema">

  <xsl:for-each select="/xml/rs:data/z:row">
  <Product>
  <xsl:for-each select="@*">
  <xsl:element name="{name()}">
  <xsl:value-of select="."/>
  </xsl:element>
  </xsl:for-each>
  </Product>
  </xsl:for-each>
  </Products>

  </xsl:template>
</xsl:stylesheet>
