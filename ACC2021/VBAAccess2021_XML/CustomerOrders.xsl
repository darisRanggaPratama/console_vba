<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" version="4.0" indent="yes"/>

<xsl:template match="dataroot">
  <html>
  <body>
  <table>
  <xsl:apply-templates select="Customers"/>
  </table>
  <table>
  <xsl:apply-templates select="Customers/Orders"/>
  </table>
  </body>
  </html>
</xsl:template>

<xsl:template match="Customers">
  <Customer>
  <ID>
  <xsl:value-of select="ID"/>
  </ID>
  <Company>
  <xsl:value-of select="Company"/>
  </Company>
  </Customer>
</xsl:template>

<xsl:template match="Customers/Orders">
  <Order>
  <OrderID>
  <xsl:value-of select="Order_x0020_ID"/>
  </OrderID>
  <CustomerID>
  <xsl:value-of select="Customer_x0020_ID"/>
  </CustomerID>
  <OrderDate>
  <xsl:value-of select="substring(Order_x0020_Date, 1, 10)"/>
  </OrderDate>
  <ShippedDate>
  <xsl:value-of select="substring(Shipped_x0020_Date, 1, 10)"/>
  </ShippedDate>
  <ShippingFee>
  <xsl:value-of select="format-number(Shipping_x0020_Fee,'####0.00')"/>
  </ShippingFee>
  </Order>
</xsl:template>

</xsl:stylesheet>
