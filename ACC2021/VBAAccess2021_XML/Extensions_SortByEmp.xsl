<?xml version="1.0"?>
    <xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:output method="xml" indent="yes"/>
      <xsl:template match="/">
      <dataroot>
      <xsl:apply-templates select="dataroot/Employees">
      <xsl:sort select="LastName" order="ascending" />
      </xsl:apply-templates>
      </dataroot>
      </xsl:template>

      <xsl:template match="//Employees">
      <Extensions>
      <FullName>
      <xsl:value-of select="LastName" />
      <xsl:text>, </xsl:text>
      <xsl:value-of select="FirstName" />
      </FullName>
      <Extension>
      <xsl:value-of select="Extension" />
      </Extension>
      </Extensions>
      </xsl:template>
    </xsl:stylesheet>
