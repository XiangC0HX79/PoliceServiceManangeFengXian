package app.controller
{
	import app.AppNotification;
	import app.model.vo.AppConfigVO;
	
	import mx.rpc.AbstractOperation;
	import mx.rpc.AsyncToken;
	import mx.rpc.events.FaultEvent;
	import mx.rpc.events.ResultEvent;
	import mx.rpc.soap.Operation;
	import mx.rpc.soap.WebService;
	import mx.utils.ObjectProxy;
	
	import org.puremvc.as3.interfaces.ICommand;
	import org.puremvc.as3.interfaces.INotification;
	import org.puremvc.as3.patterns.command.SimpleCommand;
	
	public class WebServiceCommand extends SimpleCommand implements ICommand
	{		
		override public function execute(note:INotification):void
		{			
			var webService:WebService = new WebService;
			webService.wsdl = AppConfigVO.webServiceUrl;					
			webService.loadWSDL();
			
			var arr:Array = note.getBody() as Array;
			var opreatorName:String = arr[0];
			var resultFunction:Function = arr[1];
			var args:Array = arr[2];
			var showLoading:Boolean = (arr.length > 3)?arr[3]:true;
			var resultFormat:String = (arr.length > 4)?arr[4]:"object";
			var loadingText:String = (arr.length > 5)?arr[5]:"正在查询数据，请等待系统响应...";
			
			var operation:Operation = webService.getOperation(opreatorName) as Operation;
			operation.addEventListener(ResultEvent.RESULT,onResult);
			operation.addEventListener(FaultEvent.FAULT,onFault);
			operation.arguments = args;
			operation.resultFormat = resultFormat;
			operation.send();	
							
			if(showLoading)
			{
				sendNotification(AppNotification.NOTIFY_APP_LOADINGSHOW,loadingText);
			}
			
			function onResult(event:ResultEvent):void
			{								
				if(showLoading)
				{
					sendNotification(AppNotification.NOTIFY_APP_LOADINGHIDE);
				}
				
				if(event.result == null)
					return;
				
				if(operation.resultFormat == "object")
				{
					if(event.result is ObjectProxy)
					{
						if(event.result.hasOwnProperty("Tables"))
						{
							var tables:Object = event.result.Tables;
							if(tables.hasOwnProperty("Table"))
							{
								resultFunction(tables.Table.Rows);
							}
							else if(tables.hasOwnProperty("Error"))
							{
								sendNotification(AppNotification.NOTIFY_APP_ALERTERROR,tables.Error.Rows[0]["ErrorInfo"]);
							}
						}
						else
						{
							resultFunction(event.result);
						}
					}
					else if(event.result is String)
					{
						if(event.result != "")
						{
							sendNotification(AppNotification.NOTIFY_APP_ALERTERROR,event.result);
						}
						else
						{
							resultFunction(event.result);
						}
					}
				}
				else if(operation.resultFormat == "e4x")
				{
					resultFunction(XML(event.result));
				}
			}
			
			function onFault(event:FaultEvent):void
			{
				if(showLoading)
				{
					sendNotification(AppNotification.NOTIFY_APP_LOADINGHIDE);
				}
				
				sendNotification(AppNotification.NOTIFY_APP_ALERTERROR,event.fault.faultString + "\n" + event.fault.faultDetail);
			}
		}
	}
}