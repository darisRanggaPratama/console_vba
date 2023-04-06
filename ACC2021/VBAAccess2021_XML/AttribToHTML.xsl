<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'
xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'
xmlns:rs='urn:schemas-microsoft-com:rowset'
xmlns:z='#RowsetSchema'
xmlns:html="http://www.w3.org/TR/REC-html40">

<xsl:template match="/">

<html>
<head>
<title>XML to HTML</title>

<style type="text/css">

table {
  font-family: arial, sans-serif;
  font-size:9px;
  border-collapse: collapse;
  width: 100%;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

th {background-color:#9acd32; color:black} 

tr:nth-child(even) {
  background-color: #dddddd;
}
</style>

</head>
<body>
<table width="100%" border="1">

<!-- generate table headings -->
<xsl:for-each 
select="xml/s:Schema/s:ElementType/s:AttributeType">
<th>
<xsl:value-of select="@name" />
</th>
</xsl:for-each>

<!-- loop through all data rows and get values for each column -->
<xsl:for-each select="xml/rs:data/z:row">
<tr>
<xsl:for-each select="@*">
<td>

<xsl:value-of select="."/>

</td>
</xsl:for-each>
</tr>
</xsl:for-each>
</table>
</body>
</html>
</xsl:template>
</xsl:stylesheet>

