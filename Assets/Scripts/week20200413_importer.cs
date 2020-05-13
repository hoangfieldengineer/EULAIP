using UnityEngine;
using System.Collections;
using System.IO;
using UnityEditor;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

public class week20200413_importer : AssetPostprocessor {
	/*private static readonly string filePath = "Assets/Data/week20200316_.xlsx";
	private static readonly string exportPath = "Assets/Data/week20200316_.asset";
	private static readonly string[] sheetNames = { "sheet1", };
	
	static void OnPostprocessAllAssets (string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
	{
		return;
		foreach (string asset in importedAssets) {
			if (!filePath.Equals (asset))
				continue;
				
			Entity_sheet1 data = (Entity_sheet1)AssetDatabase.LoadAssetAtPath (exportPath, typeof(Entity_sheet1));
			if (data == null) {
				data = ScriptableObject.CreateInstance<Entity_sheet1> ();
				AssetDatabase.CreateAsset ((ScriptableObject)data, exportPath);
				data.hideFlags = HideFlags.NotEditable;
			}
			
			data.sheets.Clear ();
			using (FileStream stream = File.Open (filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
				IWorkbook book = null;
				if (Path.GetExtension (filePath) == ".xls") {
					book = new HSSFWorkbook(stream);
				} else {
					book = new XSSFWorkbook(stream);
				}
				
				foreach(string sheetName in sheetNames) {
					ISheet sheet = book.GetSheet(sheetName);
					if( sheet == null ) {
						Debug.LogError("[QuestData] sheet not found:" + sheetName);
						continue;
					}

					Entity_sheet1.Sheet s = new Entity_sheet1.Sheet ();
					s.name = sheetName;
				
					for (int i=1; i<= sheet.LastRowNum; i++) {
						IRow row = sheet.GetRow (i);
						ICell cell = null;
						
						Entity_sheet1.Param p = new Entity_sheet1.Param ();
						
					cell = row.GetCell(0); p.IP = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(1); p.LicenseHash = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(2); p.MachineId = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(3); p.Version = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(4); p.Location = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(5); p.Userid = (cell == null ? 0.0 : cell.NumericCellValue);
					cell = row.GetCell(6); p.count = (int)(cell == null ? 0 : cell.NumericCellValue);
					cell = row.GetCell(7); p.SerialNumber = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(8); p.LicenseType = (cell == null ? "" : cell.StringCellValue);
					cell = row.GetCell(9); p.ShortVer = (cell == null ? "" : cell.StringCellValue);
						s.list.Add (p);
					}
					data.sheets.Add(s);
				}
			}

			ScriptableObject obj = AssetDatabase.LoadAssetAtPath (exportPath, typeof(ScriptableObject)) as ScriptableObject;
			EditorUtility.SetDirty (obj);
		}
	}*/
}
