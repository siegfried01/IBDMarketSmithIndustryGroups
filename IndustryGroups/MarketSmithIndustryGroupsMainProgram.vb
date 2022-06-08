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
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Extended Stocks (X)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">RS Line New High (H)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD Live Ready (R)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD Live Watch (W)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Long Term Leaders (L)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">RS Line Blue Dot (B)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD 50 Index (5)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">IBD Big Cap 20 (2)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Top 30 RS Rating High Vol (V)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Additions (A)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s66"><Data ss:Type="String">Deletions (D)</Data><NamedCell
                                                              ss:Name="_FilterDatabase"/></Cell>

                                                  </Row>
                                                  <Row ss:AutoFitHeight="0">
                                                      <Cell><Data ss:Type="Number">1</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
                                                      <Cell ss:StyleID="s62" ss:HRef="https://marketsmith.investors.com/mstool?Symbol=G3722&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"><Data
                                                                                                                                                                                                                                                                                  ss:Type="String">G3722</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
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
        Try
            Dim TopMemberCount = 19
            Dim filesAlreadyLoaded = New HashSet(Of String)
            Dim additionalFiles = New HashSet(Of String)
            Dim mostIndustryGroups = "MinDollarVol20MComp80.csv"
            filesAlreadyLoaded.Add(mostIndustryGroups)
            filesAlreadyLoaded.Add("197 Industry Groups.csv")
            Dim ig = IndustryGroupstToEquity.LoadTable($"%USERPROFILE%\Downloads\{mostIndustryGroups}")
            filesAlreadyLoaded.Add(mostIndustryGroups)
            Dim fileNameList = New SortedDictionary(Of String, (String, Int16)) From {
            {"Extended Stocks", ("X", 13)},
            {"RS Line New High", ("H", 14)},
            {"IBD Live Ready", ("R", 15)},
            {"IBD Live Watch", ("W", 16)},
            {"Long Term Leaders", ("L", 17)},
            {"RS Line Blue Dot", ("B", 18)},
            {"IBD 50 Index", ("5", 19)},
            {"IBD Big Cap 20", ("2", 20)},
            {"Top 30 EPS Rating Stocks with High Avg. Volume", ("V", 21)},
            {"Additions", ("A", 22)},
            {"Deletions", ("D", 23)}
            }
            '{"Large Cap", ("l", 24)},
            '{"Mid Cap", ("m", 25)},
            '{"Small Cap", ("s", 26)}


            For Each fileName In fileNameList.Keys
                If filesAlreadyLoaded.Contains(fileName) Then Continue For
                filesAlreadyLoaded.Add(fileName)
            Next

            For Each fileName In Directory.GetFiles(System.Environment.ExpandEnvironmentVariables("%DN%"), "*.csv")
                fileName = Path.GetFileName(fileName)
                If filesAlreadyLoaded.Contains(fileName) Then Continue For
                additionalFiles.Add(fileName)
            Next

            Dim marketSmithListColumnNames = New SortedDictionary(Of Int16, String)
            Dim marketSmithLists = New Dictionary(Of String, HashSet(Of String))
            Dim hrefStyle = "s62"
            Dim nsMgr = New XmlNamespaceManager(New NameTable())
            nsMgr.AddNamespace("", "urn:schemas-microsoft-com:office:spreadsheet")
            nsMgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office")
            nsMgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel")
            nsMgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")
            nsMgr.AddNamespace("html", "http://www.w3.org/TR/REC-html40")
            Dim ss As XNamespace = "urn:schemas-microsoft-com:office:spreadsheet"
            Dim x As XNamespace = "urn:schemas-microsoft-com:office:excel"
            AdjustIndustryGroupTableColumnCount(TopMemberCount + additionalFiles.Count, nsMgr, ss)
            AdjustWorksheetColumnCount(TopMemberCount + additionalFiles.Count, nsMgr, ss)
            Dim industryGroupRows As IEnumerable(Of XElement) = industryGroups.XPathSelectElements("ss:Workbook/ss:Worksheet/ss:Table/ss:Row", nsMgr)
            Dim addedColumns = AddAdditionalColumnHeaders(additionalFiles, ss, industryGroupRows, fileNameList, 24)
            AddTopMembersColumnHeaders(TopMemberCount, ss, industryGroupRows)

            For Each name In fileNameList.Keys
                marketSmithLists(name) = LoadListFromCsv.LoadListFromCsv("%USERPROFILE%\Downloads\" & name & ".csv")
                marketSmithListColumnNames(fileNameList(name).Item2) = name
            Next

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
                    Dim newCellValue = New XElement(ss + "Cell", New XAttribute(ss + "StyleID", hrefStyle), New XAttribute(ss + "HRef", $"https://marketsmith.investors.com/mstool?Symbol={industryGroupCode}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), industryGroupCode), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"), industryGroupName))
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
                                WriteLine($"count2={count2} is max (={max})--------------------------------------------")
                                row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s71"), New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count2), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                            Else
                                WriteLine($"Not a maximum : count2={count2} max = {max}")
                                row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), count2), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                            End If
                        Next

                        Dim count = 0
                        For Each e In stocksInCurrentIndustryGroup
                            'Write($" e={e} {count}/{stocksInCurrentIndustryGroup.Count}")
                            If count < TopMemberCount Then
                                Dim annotations = ""
                                Dim annoCount = 0
                                For Each name In fileNameList.Keys
                                    Dim list = marketSmithLists(name)
                                    Dim newAnnotation = fileNameList(name).Item1
                                    If newAnnotation <> "" And list.Contains(e.TickerSymbol) Then
                                        If annoCount = 0 Then
                                            annotations = "-"
                                        End If
                                        annotations = annotations & newAnnotation
                                        annoCount += 1
                                    End If
                                Next
                                row.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", hrefStyle), New XAttribute(ss + "HRef", $"https://marketsmith.investors.com/mstool?Symbol={e.TickerSymbol}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0"), New XAttribute(x + "HRefScreenTip", "comp=" & e.Composite & " RS=" & e.RS & " SMR=" & e.SMR & " $vol=" & e.DollarVolume & " EPS=" & e.EPS), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), e.TickerSymbol & annotations), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"), industryGroupName)))
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
        Catch exception As Exception
            WriteLine($"{exception.Message}")
            WriteLine($"Press any key to continue...")
            Try
                ReadKey()
            Catch ex As Exception
            End Try
        End Try

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

    Function AddAdditionalColumnHeaders(ByRef fileNameList As IEnumerable(Of String), ss As XNamespace, ByRef industryGroupRows As IEnumerable(Of XElement), ByRef loadedFiles As SortedDictionary(Of String, (String, Int16)), count As Int16) As Int16
        Dim headerRow = industryGroupRows.First()
        Dim additionalColumns = 0
        ' <Cell ss:StyleID="s66"><Data ss:Type="String">Deletions (D)</Data><NamedCell ss:Name = "_FilterDatabase" /></Cell>
        For Each fileName In fileNameList
            If fileName.EndsWith(".csv") Then
                fileName = fileName.Substring(0, fileName.Length - 4)
            End If
            If loadedFiles.ContainsKey(fileName) Then
            Else
                headerRow.Add(New XElement(ss + "Cell", New XAttribute(ss + "StyleID", "s66"), New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), fileName), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
                loadedFiles.Add(fileName, ("", count))
                count += 1
                additionalColumns += 1
            End If
        Next
        Return additionalColumns
    End Function
    Private Sub AddTopMembersColumnHeaders(ByRef TopMemberCount As Integer, ss As XNamespace, ByRef industryGroupRows As IEnumerable(Of XElement))
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
