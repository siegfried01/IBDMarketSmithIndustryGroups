Imports System
Imports System.Xml
Imports System.Xml.XPath
Imports System.Console
Module MarketSmithIndustryGroupsMainProgram
    Dim industryGroups As XDocument = <?xml version="1.0"?>
                                      <?mso-application progid="Excel.Sheet"?>
                                      <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                                          xmlns:o="urn:schemas-microsoft-com:office:office"
                                          xmlns:x="urn:schemas-microsoft-com:office:excel"
                                          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
                                          xmlns:html="http://www.w3.org/TR/REC-html40">
                                          <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
                                              <Author>Siegfried Heintze</Author>
                                              <LastAuthor>Siegfried Heintze</LastAuthor>
                                              <Created>2010-07-07T15:59:25Z</Created>
                                              <LastSaved>2022-04-09T23:57:14Z</LastSaved>
                                              <Version>16.00</Version>
                                          </DocumentProperties>
                                          <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
                                              <AllowPNG/>
                                          </OfficeDocumentSettings>
                                          <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
                                              <WindowHeight>18165</WindowHeight>
                                              <WindowWidth>19935</WindowWidth>
                                              <WindowTopX>32767</WindowTopX>
                                              <WindowTopY>32767</WindowTopY>
                                              <ProtectStructure>False</ProtectStructure>
                                              <ProtectWindows>False</ProtectWindows>
                                          </ExcelWorkbook>
                                          <Styles>
                                              <Style ss:ID="Default" ss:Name="Normal">
                                                  <Alignment ss:Vertical="Bottom"/>
                                                  <Borders/>
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
                                                  <Interior/>
                                                  <NumberFormat/>
                                                  <Protection/>
                                              </Style>
                                              <Style ss:ID="s62" ss:Name="Hyperlink">
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"
                                                      ss:Underline="Single"/>
                                              </Style>
                                              <Style ss:ID="s63" ss:Name="List Panel Header">
                                                  <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
                                                  <Borders/>
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
                                                      ss:Bold="1"/>
                                                  <Interior/>
                                                  <NumberFormat/>
                                                  <Protection/>
                                              </Style>
                                              <Style ss:ID="s64">
                                                  <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
                                                      ss:Bold="1"/>
                                              </Style>
                                              <Style ss:ID="s66" ss:Parent="s63">
                                                  <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:Rotate="90"/>
                                                  <Borders/>
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
                                                      ss:Bold="1"/>
                                                  <Interior/>
                                                  <NumberFormat/>
                                                  <Protection/>
                                              </Style>
                                          </Styles>
                                          <Worksheet ss:Name="Export">
                                              <Names>
                                                  <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=Export!R1C1:R198C34"
                                                      ss:Hidden="1"/>
                                              </Names>
                                              <Table ss:ExpandedColumnCount="34" ss:ExpandedRowCount="198" x:FullColumns="1"
                                                  x:FullRows="1" ss:DefaultColumnWidth="72" ss:DefaultRowHeight="15">
                                                  <Column ss:Width="40.5" ss:Span="1"/>
                                                  <Column ss:Index="3" ss:Width="144"/>
                                                  <Column ss:Width="40.5" ss:Span="20"/>
                                                  <Row ss:AutoFitHeight="0" ss:Height="111" ss:StyleID="s64">
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Order</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Symbol</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Name</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Number of Stocks</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Ind Group Rank</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Ind Grp Rnk Last Week</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Ind Grp Rnk 3 Mo Ago</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Ind Grp Rnk 6 Mo Ago</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">% Chg YTD</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Ind Mkt Val (bil)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Week Rank Change</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">3 Month Rank Change</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">6 Month Rank Change</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Extended Stocks</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">RS Line New High</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD Live Ready</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD Live Watch</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Long Term Leaders</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">RS Line Blue Dot</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD 50 Index</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD Big Cap 20</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Top 30 RS Rating High Vol</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Additions</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Deletions</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:MergeAcross="9"><Data ss:Type="String">Top Members</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Agricultural Operations</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">127</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">119</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8.5299999999999994</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1001</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Outsourcing</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">50</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-4.38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">250</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">139</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1002</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Office Supplies Mfg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">154</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">161</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">150</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">159</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-14.76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1004</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer-Integrated Syst</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">189</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">187</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">189</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">194</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">122</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1005</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Diversified</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-0.08</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1576</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1007</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Consumer Prod-Specialty</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">183</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">168</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">186</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1008</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-Super Regional</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">153</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">157</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">372</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1009</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Insurance-Diversified</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">80</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6.31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">520</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">191</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1010</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Trucks &amp; Parts-Hvy Duty</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">178</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">160</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">152</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">192</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-14.75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">66</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1011</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Staffing</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">113</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-6.3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">128</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1031</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Outpnt/Hm Care</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">50</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">184</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">177</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4.71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">137</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1040</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Mining-Gold/Silver/Gems</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">184</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18.14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">307</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">131</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1044</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Services</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">90</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">102</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">146</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">102</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-15.44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">154</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1094</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Consumer Elec</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">163</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">173</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">195</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">140</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-6.64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1099</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Mining-Metal Ores</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">37</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">51</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24.8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">978</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">149</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1310</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-U S Expl&amp;Prod</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34.96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">341</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">147</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1311</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Royalty Trust</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18.079999999999998</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">140</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1312</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Cdn Expl&amp;Prod</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">47.21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">144</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G1315&amp;amp;Periodicity=Daily&amp;amp;InstrumentType=Stock&amp;amp;Source=sitemarketcondition&amp;amp;AlertSubId=8241925&amp;amp;ListId=0&amp;amp;ParentId=0"><Data
                                                                                                                                                                                                                                                                                                          ss:Type="String">G1315</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Intl Expl&amp;Prod</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60.35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">299</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">143</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1317</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Integrated</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45.4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19220</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1318</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Energy-Alternative/Other</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">144</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2.59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">70</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1319</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Energy-Coal</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">63.35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">212</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1320</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Energy-Solar</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">78</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">192</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">165</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-8.5299999999999994</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">142</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1380</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Field Services</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">107</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">38.6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1381</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Drilling</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">173</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">82.86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1440</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-Foreign</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">116</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6.75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10743</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1520</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Resident/Comml</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">165</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">171</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-29.07</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">113</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1621</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Heavy Construction</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">113</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">131</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-0.08</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G1800</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Chemicals-Agricultural</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">54</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39.159999999999997</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2010</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Food-Meat Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">128</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">130</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-0.81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">105</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2020</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Food-Dairy Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">147</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">149</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">194</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">193</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43.26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">90</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2041</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Food-Grain &amp; Related</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">122</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16.82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2070</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Food-Confectionery</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">108</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8.0500000000000007</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2085</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Beverages-Alcoholic</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">66</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">108</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">112</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">167</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-8.23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">830</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2086</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Beverages-Non-Alcoholic</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.73</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">834</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">93</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2091</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Food-Packaged</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">129</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">164</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-1.57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">946</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2092</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Food-Misc Preparation</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">115</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">160</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-4.97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">183</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Tobacco</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">145</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">162</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3.07</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">359</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2300</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Apparel-Clothing Mfg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">140</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">150</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">153</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">185</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">987</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2400</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Wood Prds</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">116</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">151</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2510</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Hsehold/Office Furniture</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">170</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">181</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">174</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">179</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.920000000000002</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">150</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2621</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Paper &amp; Paper Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-1.96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2653</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Containers/Packaging</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">102</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">101</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-3.2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">123</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">118</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2711</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Media-Newspapers</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">146</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">143</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">50</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">168</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">117</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2712</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Media-Diversified</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">147</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-4.21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">557</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">119</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2721</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Media-Periodicals</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">66</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40.21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">116</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2731</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Media-Books</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">70</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">126</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">145</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2751</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Document Mgmt</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">109</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">132</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8.98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2761</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comp Sftwr-Spec Enterprs</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">191</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">185</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">170</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-24.14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2818</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Chemicals-Basic</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">50</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">87</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">109</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-0.55000000000000004</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">373</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2821</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Financial</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">136</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">154</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">271</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">123</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2830</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Ethical Drugs</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-0.4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1718</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">129</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2831</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">142</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">108</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">128</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">149</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-10.53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">911</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">172</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2840</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Soap &amp; Clng Preparatns</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">160</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">180</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">105</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">174</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-15.92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2844</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Cosmetics/Personal Care</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">119</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">133</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">709</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2851</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Chemicals-Paints</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">168</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">178</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">84</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23.04</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">118</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2899</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Chemicals-Specialty</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">93</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-12.57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">377</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">146</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G2900</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Refining/Mktg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21.91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">233</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3011</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Auto/Truck-Tires &amp; Misc</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">176</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">156</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">188</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3069</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Medical</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">174</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">193</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3079</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Chemicals-Plastics</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">123</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">103</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">156</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.08</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Apparel-Shoes &amp; Rel Mfg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">195</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">195</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">157</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-26.28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">211</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3220</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Security</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">54</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">66</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">339</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3270</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Desktop</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">186</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">189</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">179</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-19.46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2485</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3299</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Constr Prds/Misc</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">162</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">131</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-27.75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">173</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3312</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Steel-Producers</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33.42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">299</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">174</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3313</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Steel-Specialty Alloys</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">155</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">166</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23.33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">102</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3334</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Internet-Content</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">144</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">134</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">131</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-10.63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3443</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3357</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Edu/Media</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">154</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-20.86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3441</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Healthcare</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">115</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">93</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-10.34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">84</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3442</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Publ Inv Fd-Eqt</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">121</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">106</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5.61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">136</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3499</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Metal Proc &amp; Fabrication</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">68</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">94</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">121</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">300</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">112</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3522</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Machinery-Farm</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">124</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">180</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9.89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">111</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3531</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Machinery-Constr/Mining</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">132</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">135</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">110</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">195</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-11.02</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">278</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">145</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3533</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Machinery/Equip</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">37</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">140</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27.22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3537</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Machinery-Mtl Hdlg/Autmn</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">174</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">136</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">187</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">173</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-12.81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">115</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3541</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Machinery-Tools &amp; Rel</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">184</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">184</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">165</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">120</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.690000000000001</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3548</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Hand Tools</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">177</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">179</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">160</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">172</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-13.71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">181</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3552</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom-Fiber Optics</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">137</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-22.18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">159</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3559</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Internet</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">166</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">171</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">181</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">161</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-13.75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2514</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">151</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3566</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Pollution Control</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">202</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">113</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3569</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Machinery-Gen Industrial</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">137</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">140</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">73</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">78</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.850000000000001</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">363</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3574</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer-Networking</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">37</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-14.55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">258</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3575</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Design</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">139</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">147</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-21.8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">235</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3577</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Wholesale-Electronics</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">54</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">119</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">50</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3578</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer-Data Storage</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">122</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">158</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-21.03</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">145</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">51</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3580</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer-Hardware/Perip</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">151</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">172</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">109</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">116</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.079999999999998</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">398</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3582</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Database</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">152</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">154</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">118</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-15.7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">347</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3583</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Enterprse</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">194</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">193</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-29.24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1224</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3584</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer Sftwr-Gaming</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">169</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">153</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">183</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">471</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3585</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-A/C &amp; Heating Prds</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">94</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">149</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-11.95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">103</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3586</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Internet-Network Sltns</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">121</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">126</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">104</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">143</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-14.05</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">78</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3611</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Elec-Scientific/Msrng</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">166</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">105</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-26.14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3621</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Electrical-Power/Equipmt</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">131</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">118</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">70</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">231</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3631</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Hsehold-Appliances/Wares</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">187</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">188</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">161</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">186</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-20.99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3651</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Consumer Prod-Electronic</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">180</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">192</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">190</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">181</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-13.35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3662</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Elec-Misc Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">167</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">151</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">90</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">126</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-25.94</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">106</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3664</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Elec-Contract Mfg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">135</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">127</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">90</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.579999999999998</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3674</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Elec-Semiconductor Equip</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">102</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23.58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">596</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3676</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Elec-Semicondctor Fablss</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">104</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23.83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">853</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">66</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3677</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Elec-Semiconductor Mfg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">113</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3345</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">68</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3680</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Electronic-Parts</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">111</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">54</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-20.010000000000002</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">133</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3711</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Auto Manufacturers</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">73</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-12.33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2330</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3714</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Auto/Truck-Original Eqp</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">165</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">139</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">111</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">129</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">116</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3715</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Auto/Truck-Replace Parts</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">188</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">177</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">158</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">139</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-20.86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3722</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Aerospace/Defense</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">134</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">124</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8.43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">905</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3791</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Mobile/Mfg &amp; Rv</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">185</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">186</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">110</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-28.99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">133</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3831</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Systems/Equip</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">112</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">116</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">168</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.440000000000001</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">253</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">132</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3840</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Supplies</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">105</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2.58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">153</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">109</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3941</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Toys/Games/Hobby</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">107</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">119</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">130</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">135</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">107</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3949</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">190</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">191</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">172</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">142</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23.46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">171</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G3999</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Security/Sfty</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">172</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">169</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">177</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">136</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-19.03</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">188</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4010</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transportation-Rail</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">153</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-6.01</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">442</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">190</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4210</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transportation-Truck</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">120</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-28.04</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">189</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4411</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transportation-Ship</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16.7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">185</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4511</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transportation-Airline</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">181</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">171</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">183</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-12.47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">974</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">184</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4512</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transport-Air Freight</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">134</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">135</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">127</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.690000000000001</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">198</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">187</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4700</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transportation-Logistics</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">133</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">124</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23.4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">93</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4811</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom Svcs- Foreign</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">101</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">93</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12.64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7507</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">120</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4830</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Media-Radio/Tv</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">80</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">167</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3.52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">177</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4891</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom Svcs-Integrated</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">128</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">145</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">139</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-12.91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">550</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">178</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4892</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom Svcs-Wireless</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">68</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">162</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">146</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3.63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">185</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">179</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4893</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom-Cable/Satl Eqp</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">171</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">162</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">156</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">87</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23.36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">180</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4894</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom-Consumer Prods</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">51</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2846</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4895</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom-Infrastructure</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">124</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">130</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">163</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">176</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.07</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">176</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4896</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom Svcs-Cable/Satl</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">179</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">176</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">178</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">107</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.559999999999999</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">298</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">193</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4911</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Utility-Electric Power</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">68</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">134</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5.38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1456</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">194</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4920</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Utility-Gas Distribution</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">106</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">190</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13.34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4922</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Oil&amp;Gas-Transprt/Pipelne</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20.03</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">514</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">195</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4941</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Utility-Water Supply</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">93</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">192</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G4942</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Utility-Diversified</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">152</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9.1300000000000008</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">970</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">166</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5013</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail/Whlsle-Auto Parts</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">66</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">51</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1.28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">174</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">167</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5014</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail/Whlsle-Automobile</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">173</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">167</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">117</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-19.09</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">134</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5022</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Whlsle Drg/Suppl</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">170</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23.7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">104</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">197</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5040</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Wholesale-Food</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">106</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">135</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5091</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Metal Prds-Distributor</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">164</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">189</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14.1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">168</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5211</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail/Whlsle-Bldg Prds</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">111</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">614</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">170</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5313</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail/Whlsle-Office Sup</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">188</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">178</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15.83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">161</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5321</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Mail Order&amp;Direct</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">197</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">197</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">197</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">187</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-36.799999999999997</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">156</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5331</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Discount&amp;Variety</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">191</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-0.46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">130</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">160</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5342</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Leisure Products</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">145</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">164</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">164</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5391</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Specialty</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">14</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">105</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">117</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">122</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">80</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">165</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5411</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Super/Mini Mkts</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10.72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">115</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">153</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5621</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Apparel/Shoes/Acc</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">161</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">182</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">143</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-21.81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">262</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">158</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5710</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Home Furnishings</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">192</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">194</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">185</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">169</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-28.96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">163</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5812</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Restaurants</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">150</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">155</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">142</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">38</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-13.17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">477</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">157</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5912</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Drug Stores</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">84</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">113</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">163</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-6.16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">176</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">169</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G5971</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail/Whlsle-Jewelry</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">143</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">146</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6020</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-Money Center</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">103</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3334</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6021</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-Northeast</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">90</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">78</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">8</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">37</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">119</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6022</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-Southeast</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">106</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">94</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-11.41</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">111</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6023</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-Midwest</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5.3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6024</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Banks-West/Southwest</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-8.3699999999999992</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">165</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6120</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Savings &amp; Loan</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">84</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5.62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6146</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Blank Check</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">697</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">109</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7.0000000000000007E-2</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">121</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">73</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6147</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Commercial Loans</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">51</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.0500000000000007</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Consumer Loans</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">155</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">152</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">37</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-18.329999999999998</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">225</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">80</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6151</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Mrtg&amp;Rel Svc</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">142</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">144</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">103</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-15.11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">100</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6310</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Insurance-Life</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">110</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">104</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">108</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-1.07</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">494</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6320</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Insurance-Acc &amp; Health</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">56</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">128</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-10.33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">101</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6330</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Insurance-Prop/Cas/Titl</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">70</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9.35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">530</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">98</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6410</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Insurance-Brokers</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">73</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">101</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">80</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">96</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-3.97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">281</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">87</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6412</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Financial Svcs-Specialty</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">70</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">90</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">74</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">20</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-10.3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">618</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">75</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6413</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-CrdtCard/PmtPr</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">158</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">190</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">176</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">68</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.9600000000000009</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1162</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">76</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6722</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-ETF / ETN</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2885</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">117</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">112</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-6.01</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6723</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Publ Inv Fd-Bond</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">228</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">156</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">158</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">123</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-13.89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">85</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6724</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Publ Inv Fd-Glbl</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">130</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">123</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">144</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">115</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-11.25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6725</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Publ Inv Fd-Bal</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">127</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">129</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6730</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Property REIT</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">47</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-6.36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1509</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6731</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Mortgage REIT</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">118</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">111</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">136</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-8.08</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">65</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">152</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G6732</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Real Estate Dvlpmt/Ops</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">44</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-12.72</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">710</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">105</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7011</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Lodging</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">101</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">73</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">25</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-7.18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">137</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7310</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Advertising</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">117</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">133</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">123</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-14.97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">135</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7340</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Maintenance &amp; Svc</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">157</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">159</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">132</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">147</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.97</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">54</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7392</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Computer-Tech Services</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">126</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">112</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">71</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-21.69</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">713</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7394</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Leasing</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">107</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-15.79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">70</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">106</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7810</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Movies &amp; Related</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">193</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">183</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">180</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">62</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-37.909999999999997</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">209</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">108</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7900</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Services</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">46</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">84</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.0299999999999994</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">251</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">104</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7901</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Gaming/Equip</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">32</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">149</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">137</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">120</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-16.27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">331</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">110</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7903</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Leisure-Travel Booking</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">17</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">164</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">142</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-8.83</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">94</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G7950</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Funeral Svcs &amp; Rel</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">87</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">87</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">63</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-13.12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">130</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8058</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Research Eqp/Svc</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">42</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">159</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">163</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">137</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.82</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">539</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">126</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8059</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Long-term Care</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">191</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">197</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2.57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">125</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8060</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Hospitals</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">40</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">151</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-1.34</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">127</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8061</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Managed Care</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">15</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">24</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">133</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">132</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3.88</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">862</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">121</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8063</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Biomed/Biotech</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">796</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">121</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">159</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">48</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-17.239999999999998</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1479</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">124</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8064</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Medical-Generic Drugs</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">170</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">166</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">155</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-9.9700000000000006</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">54</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">78</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8072</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Investment Mgmt</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">92</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">81</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">78</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">35</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-14.3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">572</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8073</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Finance-Invest Bnk/Bkrs</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">120</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">103</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-15.22</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">379</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">19</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8074</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Bldg-Cement/Concrt/Ag</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">129</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">122</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">150</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-19.149999999999999</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">300</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">186</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8075</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Transportation-Equip Mfg</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">39</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">86</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">157</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2.0499999999999998</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">23</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">162</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8076</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Major Disc Chains</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">30</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">61</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">59</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6.06</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1518</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">155</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8077</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Retail-Department Stores</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">36</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12.27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">29</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">57</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8240</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Consumer Svcs-Education</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">45</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">114</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">110</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">169</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">118</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4.8600000000000003</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8242</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Consulting</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">53</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">94</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2.4500000000000002</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">37</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G8244</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Comml Svcs-Market Rsrch</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">13</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">115</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-11.31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">380</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">G9900</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Diversified Operations</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">28</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">77</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">104</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3.67</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3158</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                              </Table>
                                              <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
                                                  <PageSetup>
                                                      <Header x:Margin="0.3"/>
                                                      <Footer x:Margin="0.3"/>
                                                      <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
                                                  </PageSetup>
                                                  <Unsynced/>
                                                  <Print>
                                                      <ValidPrinterInfo/>
                                                      <HorizontalResolution>1200</HorizontalResolution>
                                                      <VerticalResolution>1200</VerticalResolution>
                                                  </Print>
                                                  <Selected/>
                                                  <Panes>
                                                      <Pane>
                                                          <Number>3</Number>
                                                          <ActiveCol>23</ActiveCol>
                                                      </Pane>
                                                  </Panes>
                                                  <ProtectObjects>False</ProtectObjects>
                                                  <ProtectScenarios>False</ProtectScenarios>
                                              </WorksheetOptions>
                                              <AutoFilter x:Range="R1C1:R198C34"
                                                  xmlns="urn:schemas-microsoft-com:office:excel">
                                              </AutoFilter>
                                          </Worksheet>
                                      </Workbook>


    Sub Main(args As String())
        Dim ig = IndustryGroupstToEquity.LoadTable("%USERPROFILE%\Downloads\MinDollarVol20MComp80.csv")
        Dim fileNameList = New SortedDictionary(Of String, (String, Int16)) From {
            {"Extended Stocks", ("X", 13)},
            {"RS Line New High", ("H", 14)},
            {"IBD Live Ready", ("R", 15)},
            {"IBD Live Watch", ("W", 16)},
            {"Long Term Leaders", ("L", 17)},
            {"RS Line Blue Dot", ("B", 18)},
            {"IBD 50 Index", ("5", 19)},
            {"IBD Big Cap 20", ("2", 20)},
            {"Top 30 RS Rating Stocks with High Avg. Volume", ("V", 21)},
            {"Additions", ("A", 22)},
            {"Deletions", ("D", 23)}}
        Dim columnNames = New SortedDictionary(Of Int16, String)
        Dim lists = New Dictionary(Of String, HashSet(Of String))
        Dim hrefStyle = "s62"
        Dim nsMgr = New XmlNamespaceManager(New NameTable())
        nsMgr.AddNamespace("", "urn:schemas-microsoft-com:office:spreadsheet")
        nsMgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office")
        nsMgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel")
        nsMgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")
        nsMgr.AddNamespace("html", "http://www.w3.org/TR/REC-html40")
        Dim ss As XNamespace = "urn:schemas-microsoft-com:office:spreadsheet"
        Dim groupRows As IEnumerable(Of XElement) = industryGroups.XPathSelectElements("ss:Workbook/ss:Worksheet/ss:Table/ss:Row", nsMgr)
        LoadListFromCsv.LoadIndustryGroups(industryGroups, groupRows, ss, "%USERPROFILE%\Downloads\197 Industry Groups.csv")

        For Each name In fileNameList.Keys
            lists(name) = LoadListFromCsv.LoadListFromCsv("%USERPROFILE%\Downloads\" & name & ".csv")
            columnNames(fileNameList(name).Item2) = name
        Next
        Dim rowCount = 0
        For Each row In groupRows
            Dim cells = row.XPathSelectElements("ss:Cell", nsMgr)
            Dim saveCell As XElement
            Dim cellCount = 0
            Dim cellValue As String
            Dim industryGroupName As String
            Dim industryGroupCode As String
            Dim currentRank As Integer
            Dim oneWeekAgoRank As Integer
            Dim threeMonthAgoRank As Integer
            Dim sixMonthAgoRank As Integer
            If rowCount > 0 Then
                cellValue = 0
                For Each cell In cells
                    cellValue = cell.XPathSelectElement("ss:Data", nsMgr).Value
                    Console.Write($"cellValue: {cellValue} {cellCount} ")
                    Select Case cellCount
                        Case 1
                            Write($"replacing ")
                            saveCell = cell
                            industryGroupCode = cellValue
                        Case 2
                            industryGroupName = cellValue
                        Case 4
                            currentRank = Integer.Parse(cellValue)
                        Case 5
                            oneWeekAgoRank = Integer.Parse(cellValue)
                        Case 6
                            threeMonthAgoRank = Integer.Parse(cellValue)
                        Case 7
                            sixMonthAgoRank = Integer.Parse(cellValue)
                    End Select
                    cellCount += 1
                Next
                Dim newCellValue = <Cell><Data <%= ss %> StyleID="s62" HRef="https://marketsmith.investors.com/mstool?Symbol={cellValue}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0">{cellValue}</Data><Data>{cellValue}</Data></Cell>
                newCellValue = New XElement(ss + "Cell", New XAttribute(ss + "StyleID", hrefStyle), New XAttribute(ss + "HRef", $"https://marketsmith.investors.com/mstool?Symbol={industryGroupCode}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), industryGroupCode), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"), industryGroupName))
                '    <Cell ss:StyleID=hrefStyle ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G1315&amp;amp;Periodicity=Daily&amp;amp;InstrumentType=Stock&amp;amp;Source=sitemarketcondition&amp;amp;AlertSubId=8241925&amp;amp;ListId=0&amp;amp;ParentId=0"><Data ss:Type="String">G1315</Data><NamedCell ss:Name="_FilterDatabase"/></Cell> <Cell><Data ss:Type="String">Oil&amp;Gas-Intl Expl&amp;Prod</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                saveCell.ReplaceWith(newCellValue)
                'row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s62"), New XAttribute(ss + "HRef", "https://marketsmith.investors.com/mstool?Symbol=MSFT&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), "MSFT")))
                '<Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), oneWeekAgoRank - currentRank), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), threeMonthAgoRank - currentRank), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), sixMonthAgoRank - currentRank), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                If rowCount > -1 Then
                End If
                If industryGroupName <> Nothing And ig.ContainsKey(industryGroupName) Then
                    Dim equities = ig(industryGroupName)
                    equities.Sort(Function(a As (TickerSymbol As String, comp As Double), b As (TickerSymbol As String, comp As Double))
                                      Return b.comp.CompareTo(a.comp) ' sort descending order by composite rating
                                  End Function)
                    For Each idx In columnNames.Keys
                        Dim columnName = columnNames(idx)
                        Dim list = lists(columnName)
                        Dim count2 = 0
                        For Each e In equities
                            If list.Contains(e.TickerSymbol) Then
                                count2 += 1
                            End If
                        Next
                        row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count2), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                    Next

                    Dim count = 0
                    For Each e In equities
                        Write($" e={e} {count}/{equities.Count}")
                        If count < 10 Then
                            Dim annotations = ""
                            Dim annoCount = 0
                            For Each name In fileNameList.Keys
                                Dim list = lists(name)
                                If list.Contains(e.TickerSymbol) Then
                                    If annoCount = 0 Then
                                        annotations = "-"
                                    End If
                                    annotations = annotations & fileNameList(name).Item1
                                    annoCount += 1
                                End If
                            Next
                            row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", hrefStyle), New XAttribute(ss + "HRef", $"https://marketsmith.investors.com/mstool?Symbol={e.TickerSymbol}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), e.TickerSymbol & annotations), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"), industryGroupName)))
                        Else
                            Exit For
                        End If
                        count += 1
                    Next
                Else
                    For Each idx In columnNames.Keys
                        Dim columnName = columnNames(idx)
                        Dim count = 0
                        row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                    Next
                End If
                WriteLine()
            Else
                WriteLine("header")
            End If
            rowCount = rowCount + 1
        Next

        Debug.WriteLine($"g2: {groupRows.ToString()}")
        System.IO.File.WriteAllText("..\..\..\IndustryGroups.xml", "<?xml version=""1.0""?>" & industryGroups.ToString().Replace("&amp;amp;", "&amp;"))
    End Sub


End Module
