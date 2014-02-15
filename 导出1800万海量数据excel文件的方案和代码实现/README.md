目前java框架中能够生成excel文件的的确不少，但是，能够生成大数据量的excel框架，我倒是没发现，一般数据量大了都会出现内存溢出，所以，生成大数据量的excel文件要返璞归真，用java的基础技术,IO流来实现。
   如果想用IO流来生成excel文件，必须要知道excel的文件格式内容，相当于生成html文件一样，用字符串拼接html标签保存到文本文件就可以生成一个html文件了。同理，excel文件也是可以的。怎么知道excel的文件格式呢？其实很简单，随便新建一个excel文件，双击打开，然后点击“文件”-》“另存为”，保存的类型为“xml表格”，保存之后用文本格式打开，就可以看到excel的字符串格式一览无遗了。

把下面的xml字符串复制到文本文件，然后保存为xls格式，就是一个excel文件。

<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:html="http://www.w3.org/TR/REC-html40">
<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Created>1996-12-17T01:32:42Z</Created>
  <LastSaved>2000-11-18T06:53:49Z</LastSaved>
  <Version>11.9999</Version>
</DocumentProperties>
<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <RemovePersonalInformation/>
</OfficeDocumentSettings>
<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>4530</WindowHeight>
  <WindowWidth>8505</WindowWidth>
  <WindowTopX>480</WindowTopX>
  <WindowTopY>120</WindowTopY>
  <AcceptLabelsInFormulas/>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
</ExcelWorkbook>
<Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="宋体" x:CharSet="134" ss:Size="12"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
</Styles>
<Worksheet ss:Name="Sheet1">
  <Table ss:ExpandedColumnCount="2" ss:ExpandedRowCount="2" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="14.25">
   <Column ss:AutoFitWidth="0" ss:Width="73.5"/>
   <Row>
    <Cell><Data ss:Type="String">zhangzehao</Data></Cell>
    <Cell><Data ss:Type="String">zhangzehao</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">zhangzehao</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>5</ActiveRow>
     <ActiveCol>3</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
</Worksheet>
<Worksheet ss:Name="Sheet2">
  <Table ss:ExpandedColumnCount="0" ss:ExpandedRowCount="0" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="14.25"/>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
</Worksheet>
<Worksheet ss:Name="Sheet3">
  <Table ss:ExpandedColumnCount="0" ss:ExpandedRowCount="0" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="14.25"/>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
</Worksheet>
</Workbook>

如果要生成千万级别以上的excel，除了这个关键点之外，还要控制IO流，如果有1000万记录，要迭代1000万次组装xml字符串，这样肯定占用相当大的内存，肯定内存溢出，所以，必须把组装的xml字符串分批用IO流刷新到硬盘里，如果是在web应用中，可以刷新到response中，web应用会自动把临时流保存到客户端的临时文件中，然后再一次性复制到你保存的路径。言归正传，分批刷新的话，可以迭代一批数据就flush进硬盘，同时把list，大对象赋值为空，显式调用垃圾回收器，表明要回收内存。这样的话，不管生成多大的数据量都不会出现内存溢出的，我曾经试过导出1亿的excel文件，都不会出现内存溢出，只是用了35分钟。
  当然，如果要把实现做的优雅一些，在组装xml字符串的时候，可以结合模板技术来实现，我个人喜好stringtemplate这个轻量级的框架，我给出的DEMO也是采用了模板技术生成的，当然velocity和freemarker都是可以，stringbuilder也行，呵呵。
