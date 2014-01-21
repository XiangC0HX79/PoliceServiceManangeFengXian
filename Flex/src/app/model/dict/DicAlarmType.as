package app.model.dict
{
	import flash.utils.Dictionary;
	
	import mx.collections.ArrayCollection;
	
	[Bindalbe]
	public class DicAlarmType
	{
		public static const ALL:DicAlarmType = new DicAlarmType({ID:0,PID:0,LEVEL:0,JQNAME:'所有类别'});
		
		public var id:Number;
		public var pid:Number;
		public var level:Number;
		public var label:String;
		
		public var color:Number;
		
		public function DicAlarmType(source:Object)
		{
			this.id = source.ID;
			this.pid = source.PID;
			this.level = source.LEVEL;
			this.label = source.JQNAME;
			this.color = source.COLOR;
		}
		
		public static var dict:Dictionary = new Dictionary;
		
		/*public static function get list():ArrayCollection
		{
			var arr:Array = new Array;
			for each (var item:DicPatrolType in dict)
			{
				arr.push(item);
			}			
			arr.sortOn("id",Array.NUMERIC);
			
			return new ArrayCollection(arr);
		}*/
		
		public static function get listAll():ArrayCollection
		{
			var arr:Array = new Array;
			
			arr.push(ALL);
			for each (var item:DicAlarmType in dict)
			{
				arr.push(item);
			}			
			arr.sortOn("id",Array.NUMERIC);
			
			return new ArrayCollection(arr);
		}
		
		public static function get listLevelFirst():ArrayCollection
		{
			var arr:Array = new Array;
			
			arr.push(ALL);
			for each (var item:DicAlarmType in dict)
			{
				if(item.level == 0)
					arr.push(item);
			}			
			arr.sortOn("id",Array.NUMERIC);
			
			return new ArrayCollection(arr);
		}
		
		public static function get listLevelSecond():ArrayCollection
		{
			var arr:Array = new Array;
			
			for each (var item:DicAlarmType in dict)
			{
				if(item.level == 1)
					arr.push(item);
			}			
			arr.sortOn("id",Array.NUMERIC);
			
			return new ArrayCollection(arr);
		}
	}
}