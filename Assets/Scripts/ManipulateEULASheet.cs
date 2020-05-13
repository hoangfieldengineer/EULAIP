using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Sirenix.OdinInspector;
using Sirenix.OdinInspector.Editor;
using Sirenix.Utilities;
using Sirenix.Utilities.Editor;
using UnityEditor;
using UnityEngine;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

[Serializable]
public class CountryAndReference
{
	public string country;
	public Entity_sheet_Modified referenceData;
	public bool shouldProcessThis = true;
	public MatchingWeeks matchWeeks;
}

[Serializable]
public class WeekAndOutputPath
{
	[FilePath] public string weekData;
	[FolderPath(AbsolutePath = true)] public string outputFolder;
	public bool shouldProcessThis = true;
}

[CreateAssetMenu(menuName = "Tools/Manipulate EULA Sheet")]
public class ManipulateEULASheet : /*OdinEditorWindow*/ ScriptableObject
{
    /*[MenuItem("Tools/Manipulate EULA Sheet")]
    private static void OpenWindow()
    {
        var window = GetWindow<ManipulateEULASheet>();
    
        // Nifty little trick to quickly position the window in the middle of the editor.
        window.position = GUIHelper.GetEditorWindowRect().AlignCenter(700, 700);
    }*/
    
    public List<WeekAndOutputPath> dataWeeks;
    public List<CountryAndReference> countries;

    [Button]
    public void DoAnalyseData()
    {
        AutomateForCountriesAndWeeks(dataWeeks, countries);
    }

    [Button] [InfoBox("Input Excel file, no .csv")]
    public Entity_sheet_Modified GenerateDataThatHasDescription([FilePath] string filePath)
		{
			string exportPath = filePath.Replace(Path.GetExtension(filePath), ".asset");


			Entity_sheet_Modified data = (Entity_sheet_Modified) AssetDatabase.LoadAssetAtPath(exportPath, typeof(Entity_sheet_Modified));
			if (data == null)
			{
				data = ScriptableObject.CreateInstance<Entity_sheet_Modified>();
				AssetDatabase.CreateAsset((ScriptableObject) data, exportPath);
				// data.hideFlags = HideFlags.NotEditable;
				data.name = Path.GetFileNameWithoutExtension(filePath);
			}

			if (data.sheets.Count > 0 && data.sheets[0].list.Count > 0)
			{
				Debug.Log($"{exportPath} will be overriden");
			}

			data.sheets.Clear();
			using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				IWorkbook book = null;
				if (Path.GetExtension(filePath) == ".xls")
				{
					book = new HSSFWorkbook(stream);
				}
				else
				{
					book = new XSSFWorkbook(stream);
				}

				for (int j = 0; j < book.NumberOfSheets; j++)
				{
					ISheet sheet = book.GetSheetAt(j);
				
					if (sheet == null)
					{
						Debug.LogError("[QuestData] sheet not found:");
						continue;
					}

					Entity_sheet_Modified.SheetModified s = new Entity_sheet_Modified.SheetModified();
					s.name = sheet.SheetName;

					for (int i = 1; i <= sheet.LastRowNum; i++)
					{
						IRow row = sheet.GetRow(i);
						ICell cell = null;

						ParamModified p = new ParamModified();

						int k = 0;
						cell = row.GetCell(k);
						p.Description = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.combinedGroup = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.group = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.rows = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.users = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.machines = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.F = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.H = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.IP = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.LicenseHash = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.MachineId = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.Version = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.Location = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.Userid = (cell == null ? 0.0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.count = (int) (cell == null ? 0 : cell.NumericCellValue);
						cell = row.GetCell(++k);
						p.SerialNumber = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.LicenseType = (cell == null ? "" : cell.StringCellValue);
						cell = row.GetCell(++k);
						p.ShortVer = (cell == null ? "" : cell.StringCellValue);
						
						s.list.Add(p);
					}

					data.sheets.Add(s);
				}
			}

			ScriptableObject obj = AssetDatabase.LoadAssetAtPath(exportPath, typeof(ScriptableObject)) as ScriptableObject;
			EditorUtility.SetDirty(obj);

			Debug.Log($"Generate intermediate data: {exportPath}");

			return data;
		}

    public Entity_sheet_Modified CreateSheetModifed(Entity_sheet1_Extra extra, [FolderPath] string input)
    {
	    string exportPath = input.Replace(Path.GetFileName(input), $"{extra.name}.asset"); 
	    
	    Entity_sheet_Modified data = (Entity_sheet_Modified) AssetDatabase.LoadAssetAtPath(exportPath, typeof(Entity_sheet_Modified));
	    if (data == null)
	    {
		    data = ScriptableObject.CreateInstance<Entity_sheet_Modified>();
		    AssetDatabase.CreateAsset((ScriptableObject) data, exportPath);
		    // data.hideFlags = HideFlags.NotEditable;
		    data.name = Path.GetFileNameWithoutExtension(exportPath);
	    }

	    if (data.sheets.Count > 0 && data.sheets[0].list.Count > 0)
	    {
		    Debug.Log($"{exportPath} will be overriden");
	    }

	    data.sheets.Clear();
	    Entity_sheet_Modified.SheetModified sheet = new Entity_sheet_Modified.SheetModified();
	    sheet.name = extra.name;
	    data.sheets.Add(sheet);

	    for (int i = 0; i < extra.list.Count; i++)
	    {
		    ParamModified param = new ParamModified(extra.list[i]);
		    data.sheets[0].list.Add(param);
	    }
	    
	    ScriptableObject obj = AssetDatabase.LoadAssetAtPath(exportPath, typeof(ScriptableObject)) as ScriptableObject;
	    EditorUtility.SetDirty(obj);
	    return data;
    }
    
    [Button]
    public Entity_sheet1 GenerateData([FilePath] string filePath)
    {
        string exportPath = filePath.Replace(Path.GetExtension(filePath), ".asset");
        
        
		Entity_sheet1 data = (Entity_sheet1)AssetDatabase.LoadAssetAtPath (exportPath, typeof(Entity_sheet1));
		if (data == null) {
			data = ScriptableObject.CreateInstance<Entity_sheet1> ();
			AssetDatabase.CreateAsset ((ScriptableObject)data, exportPath);
			// data.hideFlags = HideFlags.NotEditable;
            data.name = Path.GetFileNameWithoutExtension(filePath);
        }

        if (data.sheets.Count > 0 && data.sheets[0].list.Count > 0)
        {
            Debug.Log($"{exportPath} existed");
            return data;
        }
        
		data.sheets.Clear ();
		using (FileStream stream = File.Open (filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
			IWorkbook book = null;
			if (Path.GetExtension (filePath) == ".xls") {
				book = new HSSFWorkbook(stream);
			} else {
				book = new XSSFWorkbook(stream);
			}

			for (int k = 0; k < book.NumberOfSheets; k++)
			{
				ISheet sheet = book.GetSheetAt(k);
			// }
			
			// foreach(string sheetName in sheetNames) {
			// 	ISheet sheet = book.GetSheet(sheetName);
				if( sheet == null ) {
					// Debug.LogError("[QuestData] sheet not found:" + sheetName);
					continue;
				}

				Entity_sheet1.Sheet s = new Entity_sheet1.Sheet ();
				s.name = sheet.SheetName;
			
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

        Debug.Log($"Generate intermediate data: {exportPath}");
        
        return data;
    }
    
    
    [Button]
    public void FilterLocation(string value, Entity_sheet1 input, Entity_sheet1 output)
    {
        List<Entity_sheet1.Param> records = input.sheets[0].list.FindAll(x => x.Location.Equals(value));
        output.sheets[0].list = records;
    }

    [Button]
    public void SortByIP(Entity_sheet1 input, bool isAscending)
    {
        if (isAscending)
        {
            input.sheets[0].list = input.sheets[0].list.OrderBy(param => param.IP).ToList();
        }
        else
        {
            input.sheets[0].list = input.sheets[0].list.OrderByDescending(param => param.IP).ToList();
        }
    }

    [Button]
    public void GroupByIPThenSortByGroupCount(Entity_sheet1 input, bool isAscending)
    {
        if (isAscending)
        {
            input.sheets[0].list = input.sheets[0].list.GroupBy(x => x.IP).OrderBy(g => g.Count())
                .SelectMany(g => g).ToList();
        }
        else
        {
            input.sheets[0].list = input.sheets[0].list.GroupBy(x => x.IP).OrderByDescending(g => g.Count())
                .SelectMany(g => g).ToList();
        }
    }

    [Button]
    public void ExpandGroupBySearchingBasedOnUserID(Entity_sheet1 input, Entity_sheet1_Extra output)
    {
        List<Entity_sheet1.Param> list = input.sheets[0].list.Clone().ToList();
        
        List<Param_Extra> retVal = new List<Param_Extra>();

        
        List<Entity_sheet1.Param> currentGroup = new List<Entity_sheet1.Param>();
        
        currentGroup.Add(list[0]);
        list.RemoveAt(0);

        int groupNumber = 1;
        int groupCount;
        int countUniqueMachineID;
        int countUniqueUserID;
        int countLicenseF;
        int countLicenseH;
        
        while (list.Count > 0)
        {
            // if still in the same group base on IP check
            if (currentGroup[currentGroup.Count - 1].IP.Equals(list[0].IP))
            {
                currentGroup.Add(list[0]);
                list.RemoveAt(0);    
            }
            else  // this row belongs to another group based on IP check, so now let process the previous group
            {
                
                // base on UserID, find matched rows in the full list to amend to this group 
                for (int i = 0; i < currentGroup.Count; i++)
                {
                    Entity_sheet1.Param param = currentGroup[i];

                    List<Entity_sheet1.Param> temp = list.FindAll(x => x.Userid.Equals(param.Userid)).ToList();

                    for (int j = 0; j < temp.Count; j++)
                    {
                        list.Remove(temp[j]);
                    }
                    
                    currentGroup.AddRange(temp);
                }

                // statistic about this group
                groupCount = currentGroup.Count;
                countUniqueMachineID = currentGroup.Select(x => x.MachineId).Distinct().Count();
                countUniqueUserID = currentGroup.Select(x => x.Userid).Distinct().Count();
                countLicenseF = currentGroup.GroupBy(x => x.MachineId)
                    .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
                    .FindAll(x => x.licenseType.Equals("F")).Count;
                countLicenseH = currentGroup.GroupBy(x => x.MachineId)
                    .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
                    .FindAll(x => x.licenseType.Equals("H")).Count;
                
                
                // let group by ip then sort by ip count within this group
                currentGroup = currentGroup.GroupBy(x => x.IP).OrderByDescending(g => g.Count())
                    .SelectMany(g => g).ToList();
                
                // finish current group
                for (int i = 0; i < currentGroup.Count; i++)
                {
                    Param_Extra pe = new Param_Extra(currentGroup[i]);
                    pe.group = groupNumber;
                    pe.rows = groupCount;
                    pe.machines = countUniqueMachineID;
                    pe.users = countUniqueUserID;
                    pe.F = countLicenseF;
                    pe.H = countLicenseH;
                    retVal.Add(pe);
                }
                currentGroup.Clear();
                
                // next group
                if (list.Count == 0)
                {
                    break;
                }
                currentGroup.Add(list[0]);
                list.RemoveAt(0);

                ++groupNumber;
            }
        }

        if (currentGroup.Count > 0)
        {
            // statistic about this group
            groupCount = currentGroup.Count;
            countUniqueMachineID = currentGroup.Select(x => x.MachineId).Distinct().Count();
            countUniqueUserID = currentGroup.Select(x => x.Userid).Distinct().Count();
            countLicenseF = currentGroup.GroupBy(x => x.MachineId)
                .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
                .FindAll(x => x.licenseType.Equals("F")).Count;
            countLicenseH = currentGroup.GroupBy(x => x.MachineId)
                .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
                .FindAll(x => x.licenseType.Equals("H")).Count;

            // finish current group
            for (int i = 0; i < currentGroup.Count; i++)
            {
                Param_Extra pe = new Param_Extra(currentGroup[i]);
                pe.group = groupNumber;
                pe.rows = groupCount;
                pe.machines = countUniqueMachineID;
                pe.users = countUniqueUserID;
                pe.F = countLicenseF;
                pe.H = countLicenseH;
                retVal.Add(pe);
            }

            currentGroup.Clear();
        }

        output.list = retVal;
    }

    [Button]
    public void SortTheGroupsByMachineCount(Entity_sheet1_Extra input)
    {
        input.list = input.list.OrderByDescending(x => x.machines).ThenBy(x => x.group).ToList();
    }

    [Button]
    public void Export(Entity_sheet1 input, [FolderPath(AbsolutePath = true)] string path)
    {    
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("IP,LicenseHash,MachineId,Version,Location,Userid,count,SerialNumber,LicenseType,ShortVer");
        
        List<Entity_sheet1.Param> list = input.sheets[0].list;
        
        foreach (Entity_sheet1.Param param in list)
        {
            sb.AppendLine(param.ToString());
        }

        string finalPath = path.EndsWith($"{Path.DirectorySeparatorChar}") ? path : $"{path}{Path.DirectorySeparatorChar}";
        using (StreamWriter file = new StreamWriter($"{finalPath}{input.name}.csv"))
        {
            file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
        }
    }
    
    [Button]
    public void ExportSheetExtra(Entity_sheet1_Extra input, [FolderPath(AbsolutePath = true)] string path)
    {    
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("Group,Data Rows Count,Count Unique UserID,Count Unique MachineID,Count License F,Count License H,IP,LicenseHash,MachineId,Version,Location,Userid,count,SerialNumber,LicenseType,ShortVer");
        
        List<Param_Extra> list = input.list;
        
        foreach (Param_Extra param in list)
        {
            sb.AppendLine(param.ToString());
        }

        string finalPath = path.EndsWith($"{Path.DirectorySeparatorChar}") ? path : $"{path}{Path.DirectorySeparatorChar}";
        string finalFileName = $"{finalPath}{input.name}.csv";
        using (StreamWriter file = new StreamWriter(finalFileName))
        {
            file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
        }
        
        Debug.Log($"Final output: {finalFileName}");
    }
    
    [Button]
    public void ExportSheetModified(Entity_sheet_Modified input, [FolderPath(AbsolutePath = true)] string path)
    {    
	    StringBuilder sb = new StringBuilder();
	    sb.AppendLine("Description,Combined Group,Group,Data Rows Count,Count Unique UserID,Count Unique MachineID,Count License F," +
	                  "Count License H,IP,LicenseHash,MachineId,Version,Location,Userid,count,SerialNumber,LicenseType," +
	                  "ShortVer");
        
	    List<ParamModified> list = input.sheets[0].list;
        
	    foreach (ParamModified param in list)
	    {
		    sb.AppendLine(param.ToString());
	    }

	    string finalPath = path.EndsWith($"{Path.DirectorySeparatorChar}") ? path : $"{path}{Path.DirectorySeparatorChar}";
	    string finalFileName = $"{finalPath}{input.name}.csv";
	    using (StreamWriter file = new StreamWriter(finalFileName))
	    {
		    file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
	    }
        
	    Debug.Log($"Final output: {finalFileName}");
    }

    [Button]
    public void AutomateAllStepsAbove([FilePath] string filePath, CountryAndReference filterCountry, [FolderPath(AbsolutePath = true)] string outputFolder)
    {
        DateTime now = DateTime.Now;

        Entity_sheet1 data = GenerateData(filePath);
        
        Entity_sheet1 dataFiltered = ScriptableObject.CreateInstance<Entity_sheet1> ();
        dataFiltered.sheets = new List<Entity_sheet1.Sheet>();
        dataFiltered.sheets.Add(new Entity_sheet1.Sheet());
        dataFiltered.name = data.name;
        
        FilterLocation(filterCountry.country, data, dataFiltered);
        SortByIP(dataFiltered, false);
        GroupByIPThenSortByGroupCount(dataFiltered, false);
        
        Entity_sheet1_Extra dataExtra = ScriptableObject.CreateInstance<Entity_sheet1_Extra> ();
        dataExtra.name = $"{filterCountry.country}_{dataFiltered.name}";
        ExpandGroupBySearchingBasedOnUserID(dataFiltered, dataExtra);
        SortTheGroupsByMachineCount(dataExtra);

        Entity_sheet_Modified dataModified = CreateSheetModifed(dataExtra, filePath);
        dataModified.output_GroupToCSV = dataModified.output_GroupStatisticToCSV = outputFolder;
        // dataModified.referenceData = filterCountry.referenceData;
        dataModified.AutomateAllStep();

        if (filterCountry.matchWeeks != null)
        {
	        if (!filterCountry.matchWeeks.weeks.Contains(dataModified))
	        {
		        filterCountry.matchWeeks.weeks.Add(dataModified);
	        }
	        filterCountry.matchWeeks.referenceData = filterCountry.referenceData;
	        filterCountry.matchWeeks.CompareAndMatchWeeks();
	        filterCountry.matchWeeks.ExportData();
        }
        else
        {
	        dataModified.ExportGroupToCSV();
	        dataModified.ExportGroupStatisticToCSV();
        }
        
        // ExportSheetExtra(dataExtra, outputFolder);
        // ExportSheetModified(dataModified, outputFolder);

        TimeSpan ts = DateTime.Now.Subtract(now);
        Debug.Log($"Took: {ts.ToString()}");
    }

    [Button]
    public void AutomateForCountriesAndWeeks([FilePath] List<WeekAndOutputPath> dataWeeks, List<CountryAndReference> countries)
    {
        for (int i = 0; i < dataWeeks.Count; i++)
        {
	        if (!dataWeeks[i].shouldProcessThis)
	        {
		        continue;
	        }
	        
            string week = dataWeeks[i].weekData;

            for (int j = 0; j < countries.Count; j++)
            {
	            if (!countries[j].shouldProcessThis)
	            {
		            continue;
	            }
	            
                AutomateAllStepsAbove(week, countries[j], dataWeeks[i].outputFolder);
            }
        }
    }

    
}
