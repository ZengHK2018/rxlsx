package com.lipi.excel
{
	import flash.utils.ByteArray;

	/**
	 * 解析excel（xlsx)文件 的类
	 * @author lipi
	 */
	public class Excel
	{
		
		
		private var _fileByteArray:ByteArray;
		private var _sheetIndex:int;
		private var _zip:Zip;
		private var _namePathKV:Object;//名字的地址的键值对
		private var _cache:Object;
		private var _jsonCache:Object;//json缓存
		
		
		/**
		 * 
		 * @param fileByteArray xlsx文件的二进制数据
		 * @param sheetIndex excel表索引，对应excel中的工作表的标签，索引从0开始
		 */
		public function Excel(fileByteArray:ByteArray,sheetIndex:int = 0)
		{
			_fileByteArray = fileByteArray;
			_zip = new Zip(_fileByteArray);
//			_namePathKV = getNamePathKV();
			_sheetIndex = sheetIndex;
			_jsonCache = {};
			_cache = {};
		}
		
		private var _sheetArray:Array;
		
		
		
		/**
		 * 取得表名对应的表路径
		 */
		private function getNamePathKV():Object
		{
			if(_namePathKV != null) return _namePathKV;
			var returnObj:Object = {};
			var ns:Namespace;
			var $xml:XML = getFileXML("xl/workbook.xml");
			ns = $xml.namespace();
			var ns_r:Namespace = $xml.namespace("r");
			var $sheetList:XMLList = $xml.ns::sheets.ns::sheet;
			var $ridKV:Object = {};
			for each(var sheet:XML in $sheetList)
			{
				var $name:String = sheet.@name;
				var $rid:String = sheet.@ns_r::id;
				$ridKV[$rid] = $name;
			}
			
			var $relsXML:XML = getFileXML("xl/_rels/workbook.xml.rels");
			ns = $relsXML.namespace();
			var relsList:XMLList = $relsXML.ns::Relationship;
			for each(var Relationship:XML in relsList)
			{
				var $Id:String = Relationship.@Id;
				if($ridKV.hasOwnProperty($Id))
				{
					var $keyName:String = ($ridKV[$Id] as String).replace(/^\s*|\s*$/g,''); 
					var $valStr:String = Relationship.@Target;
					returnObj[$keyName] = "xl/" + $valStr;
				}
				
			}
			_namePathKV = returnObj;
			return _namePathKV;
		}
		
		//取zip中的XML文件
		private function getFileXML(url:String):XML
		{
			var $workbookUrl:String = url;//"xl/workbook.xml";
			var $byteArray:ByteArray = _zip.getFile($workbookUrl);
			var xmlString:String = $byteArray.readUTFBytes($byteArray.bytesAvailable);
			var $xml:XML = new XML(xmlString);
			return $xml;
		}
		
		/**
		 * 返回初始化指定的表数据
		 */
		public function getSheetArray():Array
		{
			return getSheetArrayUseIndex(_sheetIndex);
		}
		
		/**
		 * 取得解析后的表格数据。值为二维数组，第一维是行索引，第二维是列索引
		 */
		public function getSheetArrayUseIndex(sheetIndex:int = 0):Array
		{
			var sheetIndexString:String = (sheetIndex + 1).toString();
			var $sheetUrl:String = "xl/worksheets/sheet" + sheetIndexString + ".xml";
			if(_cache.hasOwnProperty($sheetUrl)) return _cache[$sheetUrl];
			else
			{
				_cache[$sheetUrl] = _getSheetArray($sheetUrl);
				return _cache[$sheetUrl];
			}
		}
		
		
		/**
		 * 取得解析后的表格数据,使用表名。值为二维数组，第一维是行索引，第二维是列索引
		 */
		public function getSheetArrayUseName(sheetName:String):Array
		{
			sheetName = sheetName.replace(/^\s*|\s*$/g,''); 
			var $sheetUrl:String = getNamePathKV()[sheetName];
			if($sheetUrl == null) return null;
			if(_cache.hasOwnProperty($sheetUrl)) return _cache[$sheetUrl];
			else
			{
				_cache[$sheetUrl] = _getSheetArray($sheetUrl);
				return _cache[$sheetUrl];
			}
		}
		
		
		/**
		 * 取得JSON格式，返回的是行数据，行数据为键值hash
		 */
		public function getSheetJSONUseName(sheetName:String):Array
		{
			if(_jsonCache.hasOwnProperty(sheetName)) return _jsonCache[sheetName];
			var arr:Array = getSheetArrayUseName(sheetName);
			var dataArr:Array = [];
			var titleArr:Array = arr[0];
			var i:int;
			for(i = 1;i<arr.length;i++)
			{
				var obj:Object = {};
				var cArr:Array = arr[i];
				if(cArr == null) break;
				for(var j:int = 0;j<titleArr.length;j++)
				{
					obj[titleArr[j]] = cArr[j];
				}
				dataArr.push(obj);
			}
			_jsonCache[sheetName] = dataArr;
			return dataArr;
		}
		
		
		
		private function _getSheetArray(url:String):Array
		{
			var $sheetUrl:String = url;
			var valueArray:Array = getValueArray();
			
			var $xml:XML = getFileXML($sheetUrl);
			var ns:Namespace = $xml.namespace();
			
			var excelArray:Array = [];
			var rowList:XMLList = $xml.ns::sheetData.ns::row;
			for each(var rowListItem:XML in rowList)
			{
				var rowIndex:int = int(rowListItem.@r) - 1;
				var rowArray:Array = [];
				var colList:XMLList = rowListItem.ns::c;
				for each(var colListItem:XML in colList)
				{
					var colIndex:int = ExcelUtil.getColIndex(colListItem.@r);
					var t:String = colListItem.@t;
					var v:String = colListItem.ns::v[0];
					if(t == "s")
					{
						rowArray[colIndex] = valueArray[int(v)];
					}else{
						rowArray[colIndex] = v;
					}
				}
				excelArray[rowIndex] = rowArray;
			}
			
			var mergeCellList:XMLList = $xml.ns::mergeCells.ns::mergeCell;
			for each(var mergeCellXML:XML in mergeCellList)
			{
				var ref:String = mergeCellXML.@ref;
				var refArr:Array = ref.split(":");
				var sArr:Array = ExcelUtil.colNameToPosition(refArr[0]);
				var eArr:Array = ExcelUtil.colNameToPosition(refArr[1]);
				var sValue:Object;
				if(excelArray[sArr[0]] != null)
				{
					sValue = excelArray[sArr[0]][sArr[1]];
				}
				for(var i:int = sArr[0];i<=eArr[0];i++)
				{
					for(var j:int = sArr[1];j<=eArr[1];j++)
					{
						if(excelArray[i] == null) excelArray[i] = [];
						excelArray[i][j] = sValue;
					}
				}
				
			}
			
//			_sheetArray = excelArray;
			return excelArray;
		}
		
		
		
		private function getValueArray():Array
		{
			var valueArray:Array = [];
			var $url:String = "xl/sharedStrings.xml";
			
			var $byteArray:ByteArray = _zip.getFile($url);
			$byteArray.position = 0;
			var xmlString:String = $byteArray.readUTFBytes($byteArray.bytesAvailable);
			
			var $xml:XML = new XML(xmlString);
			var ns:Namespace = $xml.namespace();
			
			var valueList:XMLList = $xml.ns::si;
			for each(var valueListItem:XML in valueList)
			{
				var textList:XMLList = valueListItem..ns::t;
				var tValue:String = "";
				for each(var textListItem:XML in textList)
				{
					var $cTValue:String = textListItem[0];
					tValue = tValue + $cTValue;
				}
				
				valueArray.push(tValue);
			}
			return valueArray;
		}
		
		
		
		
		
		
	}
}