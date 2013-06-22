rxlsx
=====

读取Excel(xlsx)文件（ActionScript版）

<pre>
private var urlloader:URLLoader;
public function RxlsxTest()
{
        urlloader = new URLLoader();
        urlloader.dataFormat = URLLoaderDataFormat.BINARY;
        urlloader.addEventListener(Event.COMPLETE,urlloaderCompHandler);
        urlloader.load(new URLRequest("Book1.xlsx"));
}
private function urlloaderCompHandler(e:Event):void
{
        var excel:Excel = new Excel(urlloader.data as ByteArray,0);
        var sheet:Array = excel.getSheetArray();//得到表格数据
}
</pre>
