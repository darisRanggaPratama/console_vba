<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" version="4.0" indent="yes"/>

<xsl:template match="dataroot">
  <html>
  <body>
  <h2 style="font-family:Verdana">Customer Orders</h2>
  <p/>
  <xsl:apply-templates select="Customers"/>
  </body>
  </html>
</xsl:template>

<xsl:template match="Customers">
<table>
  <tr>
  <td style="background-color:#FFCC33; color:#000000;">
  <xsl:value-of select="ID"/>
  </td>
  <td><b>
  <xsl:value-of select="Company"/>
  </b></td>
  </tr>
</table>
<table cellpadding="5" cellspacing="5">
  <tr style="background-color:black; color:white;">
  <td style="background-color:black; width:10px;"/>
  <td>Order ID</td>
  <td>Order Date</td>
  <td>Shipped Date</td>
  <td>Shipping Fee</td>
  </tr>
  <xsl:apply-templates select="Orders"/>
</table>
</xsl:template>

<xsl:template match="Orders">
  <tr>
  <td style="background-color:black; width:10px;"/>
  <td><xsl:value-of select="Order_x0020_ID"/></td>
  <td><xsl:value-of select="substring(Order_x0020_Date, 1, 10)"/></td>
  <td><xsl:value-of select="substring(Shipped_x0020_Date, 1, 10)"/></td>
  <td>$<xsl:value-of select="format-number(Shipping_x0020_Fee,'####0.00')"/></td>
  </tr>
</xsl:template>

</xsl:stylesheet>
