using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel; // necessary for attributes in properties
using System.Threading;
using System.Reflection;
using System.IO;
using System.Text;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using log4net;


namespace ObjectPrinterLib
{
	public class GenericListExport
	{
		[NonSerialized]
		private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
			
		#region Export 
		
		public delegate void SetCell ( int row, int column, object data, string format);
		public delegate void SetColumnWidth ( int column, int with);
		
		private class ColProp
		{
			public int column;
			public string columnName;
			
			public PropertyInfo propertyInfo;
			public FieldInfo fieldInfo;

			public string format="";
			public int width=50;
			
			public static ColProp Init<T> (string columnNameAndFormat, int columnPosition)
			{
				
				char [] colSep = new char[] {':'};
				string[] cols = columnNameAndFormat.Split(colSep);
				
				string columnName = cols[0];
				
				string colFormat = "";
				if (cols.Length>1) colFormat = cols[1];
				
				int colWidth = 0; //0 == auto width
				if (cols.Length>2) colWidth = int.Parse (cols[2]);
				
				PropertyInfo pi= typeof(T).GetProperty (columnName);
				if (pi !=null)
				{
					ColProp cp =new ColProp ();
					cp.column = columnPosition;
					cp.propertyInfo = pi;
					
					cp.columnName = columnName;
					cp.format = colFormat;
					cp.width = colWidth;
					
					return cp;
				}
				else
				{
					FieldInfo fi= typeof(T).GetField (columnName);
					if (fi !=null)
					{
						ColProp cp =new ColProp ();
						cp.column = columnPosition;
						cp.fieldInfo = fi;
						
						cp.columnName = columnName;
						cp.format = colFormat;
						cp.width = colWidth;
						return cp;
						
					}
					else
					{
						return null;
					}
				}
				
			}
			
			
			public object data (object row)
			{
				if (propertyInfo != null)
				{
					return propertyInfo.GetValue (row);
				}
				if (fieldInfo != null)
				{
					return fieldInfo.GetValue (row);
				}
				return null;
			}
		}
		
		public static void Export<T> ( List<T> list, string columnFormat, SetCell setCell, SetColumnWidth  setColumnWidth)
		{
			// Calculate List of Fields/Properties to be exported
			List<ColProp> exportColumns = new List<ColProp> ();
			char [] colSep = new char[] {';'};
			string[] cols = columnFormat.Split(colSep);
			foreach (var columnName in cols)
			{
				ColProp cp= ColProp.Init<T> ( columnName, exportColumns.Count);
				if (cp != null)
				{
					exportColumns.Add (cp);
				}
			}
			
			
			
			int row = 0;
			foreach (T t in list)
			{
				row++;
				foreach( var cp in exportColumns)
				{
					setCell (row,  cp.column, cp.data(t),  cp.format);
				}
			}
			
			foreach( var cp in exportColumns)
			{
				setCell (0,  cp.column, cp.columnName,"");
				setColumnWidth (cp.column , cp.width);
			}
			
			
			
			
		}
		
		#endregion
		
		#region Html
		
		public static string ExportHtml<T> ( List<T> list,string columnFormat)
		{
			StringBuilder sb = new StringBuilder ();
			sb.AppendLine ("<table>");
			int lastRowNumber = -1;
			bool isFirstRow = true;
			Export<T> (list, columnFormat,
			           (row, column, data,format) => {
			           	
			           	bool isNewRow = row != lastRowNumber;
			           	if (isNewRow)
			           	{
			           		if (!isFirstRow) sb.AppendLine ("</tr>");
			           		lastRowNumber=row;
			           		isFirstRow=false;
			           		sb.Append ("<tr>");
			           	}
			           	sb.Append ("<td>");
			           	sb.Append (data.ToString());
			           	sb.Append ("</td>");
			           	//if (!string.IsNullOrEmpty (format)) {
			           	//	ws.Cells[row+startRow, column+startColumn].Style.Numberformat.Format = format;
			           	//}
			           },
			           (column, width) => {
			           	//if (width ==0)
			           	//	ws.Column(column+startColumn).AutoFit();
			           	//else{
			           	//	ws.Column(column+startColumn).Width = (int) width/2;
			           	//	log.InfoFormat ("Setting Column {0} width to {1}", column,width);
			           	//}
			           }
			          );
			
			if (!isFirstRow)
			{
				sb.AppendLine("</tr>");
			}
			
			sb.AppendLine ("\r\n</table>");
			return sb.ToString ();
		}
		
		
		#endregion
		
		
		#region Excel
		
		public static void ExportExcel<T> ( List<T> list,string columnFormat,ExcelWorksheet ws, int startRow=1, int startColumn=1)
		{
			Export<T> (list, columnFormat,
			           (row, column, data,format) => {
			           	ws.Cells[row+startRow, column+startColumn].Value =data;
			           	if (!string.IsNullOrEmpty (format)) {
			           		ws.Cells[row+startRow, column+startColumn].Style.Numberformat.Format = format;
			           	}
			           },
			           (column, width) => {
			           	if (width ==0)
			           		ws.Column(column+startColumn).AutoFit();
			           	else{
			           		ws.Column(column+startColumn).Width = (int) width/2;
			           		log.InfoFormat ("Setting Column {0} width to {1}", column,width);
			           	}
			           }
			          );
		}
		
		public static void ExportExcel<T> ( List<T> list,string columnFormat, string fnExport)
		{
			ExcelPackage pck = new ExcelPackage();
			ExcelWorksheet ws = pck.Workbook.Worksheets.Add("export");

			int startRow=1;
			int startColumn=1;
			
			ExportExcel<T> (  list,columnFormat, ws, startRow, startColumn);
			
			FileInfo fi = new FileInfo(fnExport);
			ws.Calculate ();
			
			ws.View.FreezePanes(2, 1);
			pck.SaveAs(fi);
		}
		#endregion
		
		
	}
}
