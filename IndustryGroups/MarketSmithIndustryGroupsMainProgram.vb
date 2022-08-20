Imports System
Imports System.Xml
Imports System.Xml.XPath
Imports System.Console
Imports System.IO
Imports System.Diagnostics
Imports System.Xml.Linq

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
                                              <LastSaved>2022-04-17T14:09:40Z</LastSaved>
                                              <Version>16.00</Version>
                                          </DocumentProperties>
                                          <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
                                              <AllowPNG/>
                                          </OfficeDocumentSettings>
                                          <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
                                              <WindowHeight>28380</WindowHeight>
                                              <WindowWidth>32767</WindowWidth>
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
                                                  <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"/>
                                                  <Borders/>
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
                                                      ss:Bold="1"/>
                                                  <Interior/>
                                                  <NumberFormat/>
                                                  <Protection/>
                                              </Style>
                                              <Style ss:ID="s68">
                                                  <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
                                                      ss:Bold="1"/>
                                              </Style>
                                              <Style ss:ID="s71">
                                                  <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
                                              </Style>
                                              <Style ss:ID="s72" ss:Parent="s62">
                                                  <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"
                                                      ss:StrikeThrough="1" ss:Underline="Single"/>
                                              </Style>
                                          </Styles>
                                          <Worksheet ss:Name="Export">
                                              <Names>
                                                  <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=Export!R1C1:R198C40"
                                                      ss:Hidden="1"/>
                                              </Names>
                                              <Table ss:ExpandedColumnCount="40" ss:ExpandedRowCount="198" x:FullColumns="1"
                                                  x:FullRows="1" ss:DefaultColumnWidth="72" ss:DefaultRowHeight="15">
                                                  <Column ss:Width="40.5" ss:Span="1"/>
                                                  <Column ss:Index="3" ss:Width="144"/>
                                                  <Column ss:AutoFitWidth="0" ss:Width="24.75" ss:Span="4"/>
                                                  <Column ss:Index="9" ss:AutoFitWidth="0" ss:Width="35.25"/>
                                                  <Column ss:AutoFitWidth="0" ss:Width="24.75" ss:Span="3"/>
                                                  <Column ss:Index="14" ss:AutoFitWidth="0" ss:Width="19.5" ss:Span="48"/>
                                                  <Row ss:AutoFitHeight="0" ss:Height="139.5" ss:StyleID="s64">
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
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G3722&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"><Data ss:Type="String">G3722</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Aerospace/Defense</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">60</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">33</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">91</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">122</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1.1000000000000001</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">857</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">6</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">64</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">9</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=HXL&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=98 RS=82 SMR=D $vol=39025 EPS=73"><Data ss:Type="String">HXL-A</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=HEI&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=95 RS=83 SMR=B $vol=53681 EPS=89"><Data ss:Type="String">HEI-B</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=TDY&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=95 RS=74 SMR=A $vol=119107 EPS=96"><Data
                                                                                                                       ss:Type="String">TDY</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=HEIA&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=90 RS=71 SMR=B $vol=33499 EPS=91"><Data ss:Type="String">HEIA</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=LMT&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=90 RS=92 SMR=C $vol=1079522 EPS=84"><Data
                                                                                                                        ss:Type="String">LMT-ARWB</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=AJRD&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=89 RS=59 SMR=B $vol=46956 EPS=85"><Data ss:Type="String">AJRD</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=CW&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=89 RS=90 SMR=B $vol=35984 EPS=75"><Data ss:Type="String">CW-B</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=AVAV&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=87 RS=88 SMR=C $vol=32296 EPS=87"><Data ss:Type="String">AVAV</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=NOC&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=87 RS=93 SMR=C $vol=471652 EPS=69"><Data
                                                                                                                       ss:Type="String">NOC-ARWB</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62"><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=GD&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=86 RS=93 SMR=C $vol=394271 EPS=61"><Data
                                                                                                                       ss:Type="String">GD-RWB</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">172</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G2840&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"><Data
                                                                                                                                                                                                                                                                                  ss:Type="String">G2840</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Soap &amp; Clng Preparatns</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">141</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">139</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">131</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">167</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-20.170000000000002</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">95</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-10</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=CHD&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=78 RS=86 SMR=B $vol=129442 EPS=81"><Data
                                                                                                                       ss:Type="String">CHD-B</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">173</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G3312&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"><Data
                                                                                                                                                                                                                                                                                  ss:Type="String">G3312</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Steel-Producers</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">16</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">12</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">129</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">55</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31.18</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">285</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">117</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">43</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">4</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">2</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=TS&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=99 RS=96 SMR=A $vol=93069 EPS=79"><Data ss:Type="String">TS-W</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=GGB&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=98 RS=89 SMR=A $vol=75625 EPS=97"><Data ss:Type="String">GGB</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=NUE&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=98 RS=98 SMR=A $vol=525896 EPS=97"><Data
                                                                                                                       ss:Type="String">NUE-5WB</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=STLD&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=98 RS=97 SMR=A $vol=207206 EPS=98"><Data
                                                                                                                       ss:Type="String">STLD-52WV</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=X&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=96 RS=96 SMR=A $vol=530523 EPS=73"><Data
                                                                                                                       ss:Type="String">X-W</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=CLF&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=95 RS=96 SMR=A $vol=596478 EPS=92"><Data
                                                                                                                       ss:Type="String">CLF-B</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=TX&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=92 RS=80 SMR=A $vol=27550 EPS=66"><Data ss:Type="String">TX</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=MT&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=89 RS=59 SMR=A $vol=127825 EPS=77"><Data
                                                                                                                       ss:Type="String">MT</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">175</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G4811&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"><Data
                                                                                                                                                                                                                                                                                  ss:Type="String">G4811</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Telecom Svcs- Foreign</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">27</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">31</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">26</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">110</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3.89</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">7358</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">21</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">79</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=AMX&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=88 RS=92 SMR=C $vol=54782 EPS=47"><Data ss:Type="String">AMX-B</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=VIV&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=88 RS=95 SMR=C $vol=22087 EPS=68"><Data ss:Type="String">VIV</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=TU&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=86 RS=90 SMR=C $vol=46398 EPS=52"><Data ss:Type="String">TU</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=BCE&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=76 RS=85 SMR=C $vol=87602 EPS=53"><Data ss:Type="String">BCE</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=VOD&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=70 RS=46 SMR=C $vol=85627 EPS=76"><Data ss:Type="String">VOD</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">196</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G3577&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"><Data
                                                                                                                                                                                                                                                                                  ss:Type="String">G3577</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="String">Wholesale-Electronics</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">5</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">49</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">52</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">148</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">187</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">-5.03</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">11</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">3</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">99</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">138</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell><Data ss:Type="Number">0</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=NSIT&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=91 RS=83 SMR=C $vol=23322 EPS=94"><Data ss:Type="String">NSIT</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=SNX&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"
                                                          x:HRefScreenTip="comp=85 RS=54 SMR=C $vol=33229 EPS=94"><Data ss:Type="String">SNX</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
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
                                                  <LeftColumnVisible>2</LeftColumnVisible>
                                                  <Panes>
                                                      <Pane>
                                                          <Number>3</Number>
                                                          <ActiveRow>2</ActiveRow>
                                                          <ActiveCol>2</ActiveCol>
                                                          <RangeSelection>R3:R197</RangeSelection>
                                                      </Pane>
                                                  </Panes>
                                                  <ProtectObjects>False</ProtectObjects>
                                                  <ProtectScenarios>False</ProtectScenarios>
                                              </WorksheetOptions>
                                              <AutoFilter x:Range="R1C1:R198C44"
                                                  xmlns="urn:schemas-microsoft-com:office:excel">
                                              </AutoFilter>
                                          </Worksheet>
                                      </Workbook>


    Sub Main(args As String())
        Dim catchExceptions As Boolean = False
        If catchExceptions Then
            Try
                Main2()
            Catch Exception As Exception
                WriteLine($"{Exception.Message}")
                WriteLine($"Press any key to continue...")
                Try
                    ReadKey()
                Catch ex As Exception
                End Try
            End Try
        Else
            Main2()
        End If


    End Sub

    Private Sub Main2()
        Dim topMemberCount = 19 ' top N members of stocklist to show
        Dim currentExcelColumn = 13
        Dim filesAlreadyLoaded = New HashSet(Of String)
        Dim additionalNonFavoriteFiles = New HashSet(Of String)
        Dim mostIndustryGroups = "MinDollarVol1MMinPrice5.csv"
        mostIndustryGroups = "MinDollarVol20MComp80.csv"
        mostIndustryGroups = "MinDollarVol10MComp50.csv"
        filesAlreadyLoaded.Add(mostIndustryGroups)
        filesAlreadyLoaded.Add("197 Industry Groups.csv")
        Dim ig = IndustryGroupstToEquity.LoadTable($"%USERPROFILE%\Downloads\{mostIndustryGroups}")

        Dim fileNameFavoritesList = New List(Of StockList) From {
            New StockList With {.Name = "IBD Live Ready", .Attributes = New StockListAttributes With {.Annotation = "R"}},
            New StockList With {.Name = "IBD Live Watch", .Attributes = New StockListAttributes With {.Annotation = "W"}},
            New StockList With {.Name = "Long Term Leaders", .Attributes = New StockListAttributes With {.Annotation = "L"}},
            New StockList With {.Name = "Long Term Leaders Watch"},
            New StockList With {.Name = "IBD 50 Index", .Attributes = New StockListAttributes With {.Annotation = "5"}},
            New StockList With {.Name = "IBD Big Cap 20", .Attributes = New StockListAttributes With {.Annotation = "2"}},
            New StockList With {.Name = "LB Leaders"},
            New StockList With {.Name = "LB Watch"},
            New StockList With {.Name = "LB Spotlight"},
            New StockList With {.Name = "LB Sector Leaders"},
            New StockList With {.Name = "LB Top 10"},
            New StockList With {.Name = "Fundamental Strength", .Attributes = New StockListAttributes With {.Annotation = "F"}},
            New StockList With {.Name = "Technical Strength", .Attributes = New StockListAttributes With {.Annotation = "T"}},
            New StockList With {.Name = "Extended Stocks", .Attributes = New StockListAttributes With {.Annotation = "X"}},
            New StockList With {.Name = "RS Line New High", .Attributes = New StockListAttributes With {.Annotation = "h"}},
            New StockList With {.Name = "All RS Line New High", .Attributes = New StockListAttributes With {.Annotation = "H"}},
            New StockList With {.Name = "RS Line Blue Dot", .Attributes = New StockListAttributes With {.Annotation = "B"}},
            New StockList With {.Name = "RS Line 5% New High", .Attributes = New StockListAttributes With {.Annotation = "%"}},
            New StockList With {.Name = "Top 30 RS Rating Stocks with High Avg. Volume", .Attributes = New StockListAttributes With {.Annotation = "v", .DisplayName = "Top 30 RS Hi Avg Volume"}},
            New StockList With {.Name = "Top 30 EPS Rating Stocks with High Avg. Volume", .Attributes = New StockListAttributes With {.Annotation = "V", .DisplayName = "Top 30 EPS Hi Avg Volume"}},
            New StockList With {.Name = "Up on Volume", .Attributes = New StockListAttributes With {.Annotation = "U"}},
            New StockList With {.Name = "Additions", .Attributes = New StockListAttributes With {.Annotation = "A", .DisplayName = "Growth 250 Additions"}},
            New StockList With {.Name = "Deletions", .Attributes = New StockListAttributes With {.Annotation = "D", .DisplayName = "Growth 250 Deletions"}},
            New StockList With {.Name = "James P. O'Shaughnessy"},
            New StockList With {.Name = "Martin Zweig"},
            New StockList With {.Name = "Peter Lynch"},
            New StockList With {.Name = "Benjamin Graham"},
            New StockList With {.Name = "Warren Buffett"},
            New StockList With {.Name = "William J. O'Neil"},
            New StockList With {.Name = "Large Cap", .Attributes = New StockListAttributes With {.Annotation = "l"}},
            New StockList With {.Name = "Mid Cap", .Attributes = New StockListAttributes With {.Annotation = "m"}},
            New StockList With {.Name = "Small Cap", .Attributes = New StockListAttributes With {.Annotation = "s"}}
        }

        For Each stockList In fileNameFavoritesList
            stockList.Attributes.ExcelColumn = currentExcelColumn
            currentExcelColumn += 1
        Next

        'SaveToXML(fileNameFavoritesList)

        ' Stocks in Favorite Lists have a custom single character identifier that appears as a suffix for the ticker symbol and in the column header in parentheses.
        Dim fileNameFavoritesMap = New SortedDictionary(Of String, StockListAttributes)

        For Each s In fileNameFavoritesList
            fileNameFavoritesMap(s.Name) = s.Attributes
        Next

        For Each fileName In fileNameFavoritesMap.Keys
            If filesAlreadyLoaded.Contains(fileName) Then Continue For
            filesAlreadyLoaded.Add(fileName)
        Next

        For Each fileName In Directory.GetFiles(System.Environment.ExpandEnvironmentVariables("%USERPROFILE%\Downloads"), "*.csv")
            fileName = Path.GetFileName(fileName)
            If filesAlreadyLoaded.Contains(fileName) Then Continue For
            additionalNonFavoriteFiles.Add(fileName)
        Next

        Dim marketSmithListColumnNames = New SortedDictionary(Of Int16, String)
        Dim marketSmithLists = New Dictionary(Of String, HashSet(Of String))
        Dim hrefStyleNormal = "s62"
        Dim hrefStyleStrikeThru = "s72"
        Dim nsMgr = New XmlNamespaceManager(New NameTable())
        nsMgr.AddNamespace("", "urn:schemas-microsoft-com:office:spreadsheet")
        nsMgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office")
        nsMgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel")
        nsMgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")
        nsMgr.AddNamespace("html", "http://www.w3.org/TR/REC-html40")
        Dim ss As XNamespace = "urn:schemas-microsoft-com:office:spreadsheet"
        Dim x As XNamespace = "urn:schemas-microsoft-com:office:excel"
        AdjustIndustryGroupTableColumnCount(topMemberCount + additionalNonFavoriteFiles.Count, nsMgr, ss)
        AdjustWorksheetColumnCount(topMemberCount + additionalNonFavoriteFiles.Count, nsMgr, ss)
        Dim industryGroupRows As IEnumerable(Of XElement) = industryGroups.XPathSelectElements("ss:Workbook/ss:Worksheet/ss:Table/ss:Row", nsMgr)
        Dim styles As XElement = industryGroups.XPathSelectElement("ss:Workbook/ss:Styles", nsMgr)
        Dim styleNames As HashSet(Of String) = New HashSet(Of String)
        'Dim filterDimensions = industryGroups.XPathEvaluate("ss:Workbook/ss:Worksheet/ss:Names/ss:NamedRange/@ss:RefersTo", nsMgr)


        For Each name In fileNameFavoritesMap.Keys
            Try
                marketSmithLists(name) = LoadListFromCsv.LoadListFromCsv("%USERPROFILE%\Downloads\" & name & ".csv")
                fileNameFavoritesMap(name).CsvFileFoundAndLoaded = True
                marketSmithListColumnNames(fileNameFavoritesMap(name).ExcelColumn) = name
            Catch ex As MissingFile
                fileNameFavoritesMap(name).CsvFileFoundAndLoaded = False
                WriteLine("Favorite Market Smith List """ & name & """ is missing")
            End Try
            ' add XML column headers here
        Next

        AddStockListExcelColumnHeaders(additionalNonFavoriteFiles, fileNameFavoritesList, fileNameFavoritesMap, currentExcelColumn, ss, industryGroupRows)        'fileNameFavoritesMap now containes all the files that were loaded and the column positions for both the favorite and non-favorite files.

        For Each name In fileNameFavoritesMap.Keys
            If marketSmithLists.ContainsKey(name) Then
                WriteLine($"{name} is already loaded")
                marketSmithListColumnNames(fileNameFavoritesMap(name).ExcelColumn) = name
            Else
                Try
                    Dim listOfStocks = LoadListFromCsv.LoadListFromCsv("%USERPROFILE%\Downloads\" & name & ".csv")
                    WriteLine($"Loading list {name}")
                    marketSmithLists(name) = listOfStocks
                    marketSmithListColumnNames(fileNameFavoritesMap(name).ExcelColumn) = name
                Catch ex As MissingFile
                    marketSmithLists.Remove(name)
                    WriteLine("Favorite Market Smith List """ & name & """ is missing")
                End Try

            End If
        Next

        AddTopMembersExcelColumnHeaders(topMemberCount, ss, industryGroupRows)

        ' compute maximum values for each column for highlighting later
        Dim marketSmithByIndustryGroupOfCount As AutoAddDictionary(Of String, AutoAddDictionary(Of String, Int32)) = IndustryGroupMarketSmithListFindMinMax(ig, marketSmithListColumnNames, marketSmithLists, nsMgr, industryGroupRows)
        Dim MarketSmithListMax As AutoAddDictionary(Of String, Int32) = New AutoAddDictionary(Of String, Int32)
        For Each idx In marketSmithListColumnNames.Keys
            Dim marketSmithColumnName = marketSmithListColumnNames(idx)
            Dim max = Int32.MinValue
            For Each industryGroupKey In marketSmithByIndustryGroupOfCount.Keys
                If marketSmithByIndustryGroupOfCount(industryGroupKey)(marketSmithColumnName) > max Then
                    max = marketSmithByIndustryGroupOfCount(industryGroupKey)(marketSmithColumnName)
                End If
            Next
            MarketSmithListMax(marketSmithColumnName) = max
        Next
        Dim rowCount = 0
        For Each row In industryGroupRows ' second pass thru industry group table
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
            If rowCount > 0 Then ' skip the header row
                cellValue = 0
                For Each cell In cells
                    cellValue = cell.XPathSelectElement("ss:Data", nsMgr).Value
                    Select Case cellCount
                        Case 1
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
                'Dim newCellValue = <Cell><Data <%= ss %> StyleID="s62" HRef="https://marketsmith.investors.com/mstool?Symbol={cellValue}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0">{cellValue}</Data><Data>{cellValue}</Data></Cell>
                Dim newCellValue = New XElement(ss + "Cell", New XAttribute(ss + "StyleID", hrefStyleNormal), New XAttribute(ss + "HRef", $"https://marketsmith.investors.com/mstool?Symbol={industryGroupCode}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), industryGroupCode), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"), industryGroupName))
                '    <Cell ss:StyleID=hrefStyle ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G1315&amp;amp;Periodicity=Daily&amp;amp;InstrumentType=Stock&amp;amp;Source=sitemarketcondition&amp;amp;AlertSubId=8241925&amp;amp;ListId=0&amp;amp;ParentId=0"><Data ss:Type="String">G1315</Data><NamedCell ss:Name="_FilterDatabase"/></Cell> <Cell><Data ss:Type="String">Oil&amp;Gas-Intl Expl&amp;Prod</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                saveCell.ReplaceWith(newCellValue)
                'row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s62"), New XAttribute(ss + "HRef", "https://marketsmith.investors.com/mstool?Symbol=MSFT&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), "MSFT")))
                '<Cell><Data ss:Type="Number">58</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), oneWeekAgoRank - currentRank), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), threeMonthAgoRank - currentRank), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), sixMonthAgoRank - currentRank), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                If industryGroupName <> Nothing And ig.ContainsKey(industryGroupName) Then
                    Dim stocksInCurrentIndustryGroup = ig(industryGroupName)
                    stocksInCurrentIndustryGroup.Sort(Function(a As Equity, b As Equity)
                                                          Return b.Composite.CompareTo(a.Composite) ' sort descending order by composite rating
                                                      End Function)
                    Dim c = marketSmithListColumnNames.Keys.Count
                    ' add counts of stocks in current list
                    For Each idx In marketSmithListColumnNames.Keys
                        Dim marketSmithColumnName = marketSmithListColumnNames(idx)
                        Dim marketSmithList = marketSmithLists(marketSmithColumnName)
                        Dim count2 = 0
                        For Each stock In stocksInCurrentIndustryGroup
                            If marketSmithList.Contains(stock.TickerSymbol) Then
                                count2 += 1
                            End If
                        Next
                        Dim max = MarketSmithListMax(marketSmithColumnName)
                        If max = count2 And max > 0 Then
                            'WriteLine($"count2={count2} is max (={max})--------------------------------------------")
                            row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s71"), New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count2), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase")))) ' this is the maximum for this column (stock list)
                        Else
                            'WriteLine($"Not a maximum : count2={count2} max = {max}")
                            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count2), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                        End If
                    Next

                    ' add ticker symbol links for top stocks in current industry group
                    Dim count = 0
                    For Each stock In stocksInCurrentIndustryGroup
                        'Write($" e={e} {count}/{stocksInCurrentIndustryGroup.Count}")
                        If count <= topMemberCount Then
                            Dim annotations = ""
                            Dim annoCount = 0
                            For Each name In fileNameFavoritesMap.Keys
                                If marketSmithLists.ContainsKey(name) Then
                                    Dim list = marketSmithLists(name)
                                    Dim newAnnotation = fileNameFavoritesMap(name).Annotation
                                    If newAnnotation <> "" And list.Contains(stock.TickerSymbol) Then
                                        If annoCount = 0 Then
                                            annotations = "-"
                                        End If
                                        annotations = annotations & newAnnotation
                                        annoCount += 1
                                    End If
                                Else
                                    ' fileNameFavoritesMap.Reverse(name)
                                End If
                            Next
                            Dim hrefStyle As String
                            If annotations.Contains("X") Then
                                hrefStyle = hrefStyleStrikeThru
                            Else
                                hrefStyle = hrefStyleNormal
                            End If
                            ' x= 40 to 100, y= 0 to 140
                            ' emacs calc: solve([40*m+b=0, 100*m+b=140],[m,b])==[m = 2.33333333333, b = -93.3333333333]
                            Dim excelStyle = New ExcelStyle() With {.HueOffset = -93.33F, .HueScale = 2.3333F, .Saturation = 200.0F, .Luminesence = 200.0F}
                            excelStyle.Hue = Convert.ToInt32(stock.Composite + 0.5)
                            If annotations.Contains("X") Then excelStyle.Font = 1
                            If stock.DollarVolume > 20 * 1000 Then
                                excelStyle.Shade = 0
                            ElseIf stock.DollarVolume > 15 * 1000 Then
                                excelStyle.Shade = 1
                            ElseIf stock.DollarVolume > 10 * 1000 Then
                                excelStyle.Shade = 2
                            ElseIf stock.DollarVolume > 5 * 1000 Then
                                excelStyle.Shade = 3
                            Else
                                excelStyle.Shade = 4
                            End If
                            Dim fonts As XElement() = {New XElement(ss + "Font", New XAttribute(ss + "FontName", "Calibri"), New XAttribute(ss + "Size", 11), New XAttribute(ss + "Bold", 1), New XAttribute(ss + "Color", "#0066CC"), New XAttribute(ss + "Underline", "Single")), New XElement(ss + "Font", New XAttribute(ss + "FontName", "Calibri"), New XAttribute(ss + "Size", 11), New XAttribute(ss + "Bold", 1), New XAttribute(ss + "Color", "#0066CC"), New XAttribute(ss + "Underline", "Single"), New XAttribute(ss + "StrikeThrough", "1"))}
                            Dim font As XElement
                            If annotations.Contains("X") Then
                                font = fonts(1)
                            Else
                                font = fonts(0)
                            End If
                            If styleNames.Contains(excelStyle.ToString()) Then
                            Else
                                styleNames.Add(excelStyle.ToString())
                                Dim border As XElement() = {New XElement(ss + "Border", New XAttribute(ss + "Position", "Bottom"), New XAttribute(ss + "LineStyle", "Continuous"), New XAttribute(ss + "Weight", "1")), New XElement(ss + "Border", New XAttribute(ss + "Position", "Left"), New XAttribute(ss + "LineStyle", "Continuous"), New XAttribute(ss + "Weight", "1")), New XElement(ss + "Border", New XAttribute(ss + "Position", "Right"), New XAttribute(ss + "LineStyle", "Continuous"), New XAttribute(ss + "Weight", "1")), New XElement(ss + "Border", New XAttribute(ss + "Position", "Top"), New XAttribute(ss + "LineStyle", "Continuous"), New XAttribute(ss + "Weight", "1"))}
                                Dim pattern As XAttribute() = {New XAttribute(ss + "Pattern", "Solid"), New XAttribute(ss + "Pattern", "Gray0625"), New XAttribute(ss + "Pattern", "ThinHorzStripe"), New XAttribute(ss + "Pattern", "ThinVertStripe"), New XAttribute(ss + "Pattern", "ThinVertStripe"), New XAttribute(ss + "Pattern", "HorzStripe"), New XAttribute(ss + "Pattern", "VertStripe")}
                                styles.Add(New XElement(ss + "Style", New XAttribute(ss + "ID", excelStyle.ToString()), New XElement(ss + "Borders", border(0), border(1), border(2), border(3)), font, New XElement(ss + "Interior", New XAttribute(ss + "Color", "#" & excelStyle.Color), pattern(excelStyle.Shade))))

                            End If
                            hrefStyle = excelStyle.ToString()
                            row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", hrefStyle), New XAttribute(ss + "HRef", $"https://marketsmith.investors.com/mstool?Symbol={stock.TickerSymbol}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XAttribute(x + "HRefScreenTip", stock.Name & ": sym=" & stock.TickerSymbol & " " & "comp=" & stock.Composite & " RS=" & stock.RS & " SMR=" & stock.SMR & " $vol=" & stock.DollarVolume & " EPS=" & stock.EPS & " Up/Down=" & stock.UpDown), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), stock.TickerSymbol & annotations), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"), industryGroupName)))
                        Else
                            Exit For
                        End If
                        count += 1
                    Next
                Else
                    For Each idx In marketSmithListColumnNames.Keys
                        Dim columnName = marketSmithListColumnNames(idx)
                        Dim count = 0
                        row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                    Next
                End If
                'WriteLine()
            Else
                'WriteLine("header")
            End If
            rowCount = rowCount + 1
        Next

        'Debug.WriteLine($"g2: {industryGroupRows.ToString()}")
        Dim outputFileName = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Downloads\IndustryGroups.xml")
        Dim excel = Environment.ExpandEnvironmentVariables("%MSOFFICE%\EXCEL.EXE")
        File.WriteAllText(outputFileName, "<?xml version=""1.0""?>" & industryGroups.ToString().Replace("&amp;amp;", "&amp;"))
        System.Diagnostics.Process.Start(excel, $"/s ""{outputFileName}""")

        'Dim result = Shell(excel & " " & outputFileName, AppWinStyle.NormalFocus, True)
    End Sub

    Private Sub SaveToXML(fileNameFavoritesList As List(Of StockList))
        Dim xmlSerializer As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(fileNameFavoritesList.GetType)
        Dim string_writer As New System.IO.StringWriter
        Dim ns As New System.Xml.Serialization.XmlSerializerNamespaces
        xmlSerializer.Serialize(string_writer, fileNameFavoritesList, ns)
        System.IO.File.WriteAllText(System.Environment.ExpandEnvironmentVariables($"%USERPROFILE%\Downloads\MarketSmithFavoriteStockLists.xml"), string_writer.ToString)
    End Sub

    Private Sub AddStockListExcelColumnHeaders(ByRef additionalNonFavoriteFiles As HashSet(Of String), ByRef fileNameFavoritesList As List(Of StockList), ByRef fileNameFavoritesMap As SortedDictionary(Of String, StockListAttributes), ByRef nextColumn As Integer, ss As XNamespace, ByRef industryGroupRows As IEnumerable(Of XElement))
        Dim headerRow = industryGroupRows.First()
        AddExcelColumnHeadersForFavoriteStockLists(fileNameFavoritesList, fileNameFavoritesMap, ss, headerRow)
        Dim addedColumns = AddRemainingExcelColumnHeaders(additionalNonFavoriteFiles, ss, industryGroupRows, fileNameFavoritesMap, fileNameFavoritesList, nextColumn, headerRow) ' load the non-favorite ones second
    End Sub

    Private Sub AddExcelColumnHeadersForFavoriteStockLists(fileNameFavoritesList As List(Of StockList), fileNameFavoritesMap As SortedDictionary(Of String, StockListAttributes), ss As XNamespace, headerRow As XElement)
        For Each stockList In fileNameFavoritesList
            If fileNameFavoritesMap.ContainsKey((stockList.Name)) Then
                If fileNameFavoritesMap(stockList.Name).CsvFileFoundAndLoaded Then
                    Dim annotation = ""
                    If fileNameFavoritesMap(stockList.Name).Annotation <> "" Then
                        annotation = " (" & fileNameFavoritesMap(stockList.Name).Annotation & ")"
                    End If
                    Dim stockListName As String
                    If String.IsNullOrEmpty(stockList.Attributes.DisplayName) Then
                        stockListName = stockList.Name
                    Else
                        stockListName = stockList.Attributes.DisplayName
                    End If
                    headerRow.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s66"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), stockListName & annotation), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                Else
                    WriteLine($"{stockList.Name} file is missing: skipping this column header")
                End If
            Else
                WriteLine($"Internal error: {stockList.Name} not found")
            End If
        Next
    End Sub

    Private Function IndustryGroupMarketSmithListFindMinMax(ig As Dictionary(Of String, List(Of Equity)), marketSmithListColumnNames As SortedDictionary(Of Short, String), marketSmithLists As Dictionary(Of String, HashSet(Of String)), nsMgr As XmlNamespaceManager, industryGroupRows As IEnumerable(Of XElement)) As AutoAddDictionary(Of String, AutoAddDictionary(Of String, Int32))
        Dim rowCount0 = 0
        Dim marketSmithByIndustryGroupOfCount = New AutoAddDictionary(Of String, AutoAddDictionary(Of String, Int32))
        For Each row In industryGroupRows 'first pass thru industry group table
            Dim cells = row.XPathSelectElements("ss:Cell", nsMgr)
            Dim cellCount = 0
            Dim cellValue As String
            Dim industryGroupName As String
            If rowCount0 > 0 Then ' skip the header row
                cellValue = 0
                For Each cell In cells
                    cellValue = cell.XPathSelectElement("ss:Data", nsMgr).Value
                    Select Case cellCount
                        Case 2
                            industryGroupName = cellValue
                            Exit For
                    End Select
                    cellCount += 1
                Next
                If industryGroupName <> Nothing And ig.ContainsKey(industryGroupName) Then
                    Dim stocksInCurrentIndustryGroup = ig(industryGroupName)
                    stocksInCurrentIndustryGroup.Sort(Function(a As Equity, b As Equity)
                                                          Return b.Composite.CompareTo(a.Composite) ' sort descending order by composite rating
                                                      End Function)
                    For Each idx In marketSmithListColumnNames.Keys
                        Dim columnName = marketSmithListColumnNames(idx)
                        Dim marketSmithList = marketSmithLists(columnName)
                        Dim count2 = 0
                        For Each stock In stocksInCurrentIndustryGroup
                            If marketSmithList.Contains(stock.TickerSymbol) Then
                                count2 += 1
                            End If
                        Next
                        marketSmithByIndustryGroupOfCount(industryGroupName)(columnName) = count2
                    Next
                End If
            End If
            rowCount0 += 1
        Next

        Return marketSmithByIndustryGroupOfCount
    End Function

    Function AddRemainingExcelColumnHeaders(ByRef fileNameList As IEnumerable(Of String), ss As XNamespace, ByRef industryGroupRows As IEnumerable(Of XElement), ByRef fileNameFavoritesMap As SortedDictionary(Of String, StockListAttributes), ByRef fileNameFavoritesList As List(Of StockList), ByRef nextColumn As Int16, headerRow As XElement) As Int16
        Dim additionalColumns = 0
        For Each fileName In fileNameList
            If fileName.EndsWith(".csv") Then
                fileName = fileName.Substring(0, fileName.Length - 4)
            End If
            If fileNameFavoritesMap.ContainsKey(fileName) Then
                'headerRow.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s66"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), fileName & " (" & fileNameFavoritesMap(fileName).Item1 & ")"), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            Else
                headerRow.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s66"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), fileName), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                fileNameFavoritesMap.Add(fileName, New StockListAttributes With {.ExcelColumn = nextColumn, .DisplayName = "", .CsvFileFoundAndLoaded = True}) ' add non-favorite to map consequetively
                nextColumn += 1
                additionalColumns += 1
            End If
        Next
        Return additionalColumns
    End Function

    Private Sub AddTopMembersExcelColumnHeaders(ByRef TopMemberCount As Integer, ss As XNamespace, ByRef industryGroupRows As IEnumerable(Of XElement))
        Dim headerRow = industryGroupRows.First()
        '
        '<Cell ss:MergeAcross="19" ss:StyleID="s68"><Data ss:Type="String">Top Members</Data><NamedCell ss:Name = "_FilterDatabase" /></Cell>
        headerRow.Add(New XElement(ss + "Cell", New XAttribute(ss + "MergeAcross", TopMemberCount.ToString), New XAttribute(ss + "StyleID", "s68"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), "Top Members"), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
        LoadListFromCsv.LoadIndustryGroups(industryGroups, industryGroupRows, ss, "%USERPROFILE%\Downloads\197 Industry Groups.csv")
    End Sub

    Private Sub AdjustWorksheetColumnCount(TopMemberCount As Integer, nsMgr As XmlNamespaceManager, ss As XNamespace)
        Dim industryGroupWorksheetNames As XElement = industryGroups.XPathSelectElements("ss:Workbook/ss:Worksheet/ss:Names/ss:NamedRange", nsMgr).First
        industryGroupWorksheetNames.SetAttributeValue(ss + "RefersTo", "Export!R1C1:R198C" & (25 + TopMemberCount).ToString)
    End Sub

    Sub AdjustIndustryGroupTableColumnCount(TopMemberCount As Integer, nsMgr As XmlNamespaceManager, ss As XNamespace)
        Dim industryGroupTable = industryGroups.XPathSelectElements("ss:Workbook/ss:Worksheet/ss:Table", nsMgr).First
        industryGroupTable.Attributes(ss + "ExpandedColumnCount").Remove
        industryGroupTable.SetAttributeValue(ss + "ExpandedColumnCount", (25 + TopMemberCount).ToString())
    End Sub
End Module
