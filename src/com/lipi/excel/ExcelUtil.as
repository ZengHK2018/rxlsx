package com.lipi.excel
{
	

	public class ExcelUtil
	{
		public function ExcelUtil()
		{
		}
		
		
		/**
		 * excel列标题对应的索引（A1，A2）
		 * @param colName
		 * @return 
		 */
		public static function getColIndex(colName:String):int
		{
			var abc:String = colName.replace(/(\d)/g,"");
			return textToInt(abc);
		}
		
		
		public static function colNameToPosition(colName:String):Array
		{
			var colText:String = colName.replace(/(\d)/g,"");
			var col:int = textToInt(colText);
			var row:int = int(colName.replace(colText,"")) - 1;
			return [row,col];
		}
		
		public static function textToInt(abc:String):int
		{
			var returnValue:int = 0;
			var maxLen:int = abc.length
			for (var i:int = 0; i < maxLen; i++) {
				var cValue:int = abc.charCodeAt(i) - 64;
				returnValue*=26
				returnValue += cValue
			}
			return returnValue-1;
		}
		
	}
}