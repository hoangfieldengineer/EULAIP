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
using System.Runtime.InteropServices;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;


[CreateAssetMenu(menuName = "ScriptableObjects/Entity Sheet Modified")]
public class Entity_sheet_Modified : ScriptableObject
{	
	public List<SheetModified> sheets = new List<SheetModified> ();
	public List<Group> groups = new List<Group>();
	
	[System.SerializableAttribute]
	public class SheetModified
	{
		public string name = string.Empty;
		[TableList(ShowIndexLabels = true, MaxScrollViewHeight = 1000, MinScrollViewHeight = 1000)]
		public List<ParamModified> list = new List<ParamModified>();
	}
	
	[HorizontalGroup("Group 1")]
	[InfoBox("Step 1")]
	[Button]
	public void ClearGroups()
	{
		groups.Clear();
	}
	
	[HorizontalGroup("Group 1")]
	[InfoBox("Step 2")]
	[Button]
	public void SheetToGroup()
	{
		List<ParamModified> list = sheets[0].list;

		if (list.Count == 0)
		{
			Debug.LogWarning("Nothing in list to process");
			return;
		}
		
		Group group = new Group();
		
		ParamModified previous = list[0];
		group.rows.Add(previous);
		group.name = previous.Description;
		group.users = previous.users;
		group.machines = previous.machines;
		group.F = previous.F;
		group.H = previous.H;
		Debug.Log($"New group: {group.name}");
		
		for (int i = 1; i < list.Count; i++)
		{
			ParamModified temp = list[i];
			
			if (temp.group != previous.group)
			{
				if (!string.IsNullOrEmpty(previous.combinedGroup) && previous.combinedGroup.Equals(temp.combinedGroup))
				{
					Debug.Log($"Merged group: {group.name}");
					group.users += temp.users;
					group.machines += temp.machines;
					group.F += temp.F;
					group.H += temp.H;
					
					if (!string.IsNullOrEmpty(temp.Description))
					{
						group.description += $"{temp.Description}\n";
					}	
				}
				else
				{
					groups.Add(group);
					group = new Group();
					group.users = temp.users;
					group.machines = temp.machines;
					group.F = temp.F;
					group.H = temp.H;
					
					if (!string.IsNullOrEmpty(temp.Description))
					{
						group.name = temp.Description;
						Debug.Log($"New group: {group.name}");
					}
				}
			}
			else
			{
				if (!string.IsNullOrEmpty(temp.Description))
				{
					group.description += $"{temp.Description}\n";
				}
			
			}
			group.rows.Add(temp);
			previous = temp;
		}
		
		group.CalculateStatistic();
		groups.Add(group);
	}
	
	[InfoBox("Step 3")]
	[Button]
	public void CorrectGroupNameAndDescriptionBetweenSheets(Entity_sheet_Modified other, float matchingThreshold = 0.3f, bool applyCorrection = false)
	{
		StringBuilder log = new StringBuilder();
		List<Group> namedGroups1 = groups.Where(x => !string.IsNullOrEmpty(x.name) || !string.IsNullOrEmpty(x.description)).ToList();
		List<Group> namedGroups2 = other.groups.Where(x => !string.IsNullOrEmpty(x.name) || !string.IsNullOrEmpty(x.description)).ToList();

		for (int i = 0; i < namedGroups1.Count; i++)
		{
			for (int j = 0; j < namedGroups2.Count; j++)
			{
				Group g1 = namedGroups1[i];
				Group g2 = namedGroups2[j];

				if (g1.Compare(g2, name, other.name, log, matchingThreshold) > matchingThreshold)
				{
					if (applyCorrection)
					{
						Debug.Log(log);
						Group groupToChoose = g1.description.Length > g2.description.Length ? g1 : g2;
						g1.name = groupToChoose.name;
						g2.name = groupToChoose.name;
						g1.description = groupToChoose.description;
						g2.description = groupToChoose.description;
					}
					else
					{
						Debug.Log(log);	
					}
				}
			}
		}
	}
	
	[InfoBox("Step 4")]
	[Button] 
	public int MatchingNamedGroupsBetweenSheets(/*Entity_sheet_Modified other,*/ float matchingThreshold = 0.2f, bool applyLabeling = false)
	{
		Entity_sheet_Modified other = referenceData;
		List<Group> namedGroups1 = groups.Where(x => !string.IsNullOrEmpty(x.name) || !string.IsNullOrEmpty(x.description)).ToList();
		List<Group> namedGroups2 = other.groups.Where(x => !string.IsNullOrEmpty(x.name) || !string.IsNullOrEmpty(x.description)).ToList();
		
		CrossCheck(namedGroups1, this.name, other.groups, other.name, matchingThreshold, applyLabeling);
		CrossCheck(namedGroups2, other.name, groups, this.name, matchingThreshold, applyLabeling);
		
		MergeDuplicatedGroupsByName();
		return other.MergeDuplicatedGroupsByName();
	}
	
	[HorizontalGroup("Group 7")]
	[InfoBox("Step 5")]
	[Button]
	public int MergeDuplicatedGroupsByName()
	{
		int retVal = 0;
		Debug.Log($"====================Merging {name}=====================\n\n\n");
		List<Group> namedGroups = groups.Where(x => !string.IsNullOrEmpty(x.name)).ToList();
		for (int i = 0; i < namedGroups.Count - 1; i++)
		{
			Group groupToCheck = namedGroups[i];
			
			for (int j = i + 1; j < namedGroups.Count; j++)
			{
				Group groupToBeMerge = namedGroups[j];
				if (groupToBeMerge.rows.Count > 0 && groupToCheck.name.Equals(groupToBeMerge.name))
				{
					string a = groupToCheck.ToString();
					string b = groupToBeMerge.ToString();
					
					groupToCheck.rows.AddRange(groupToBeMerge.rows);
					// groupToCheck.F += groupToBeMerge.F;
					// groupToCheck.H += groupToBeMerge.H;
					// groupToCheck.machines += groupToBeMerge.machines;
					// groupToCheck.users += groupToBeMerge.users;
					groupToBeMerge.rows.Clear();
					groupToCheck.CalculateStatistic();
					Debug.Log($"Merge [{groupToCheck.name}]\n{b}\ninto\n{a}\n=> {groupToCheck}");
					
					groups.Remove(groupToBeMerge);

					++retVal;
				}
			}
		}
		Debug.Log($"==================================================================================\n\n\n");
		
		AssignCombinedGroupColumn();
		CalculateGroupStatistic();
		GroupToSheet();

		return retVal;
	}
	
	[HorizontalGroup("Group 7")]
	[InfoBox("Step 6")]
	[Button]
	public void AssignCombinedGroupColumn()
	{
		for (int i = 0; i < groups.Count; i++)
		{
			List<int> distinctGroup = groups[i].rows.Select(x => x.group).Distinct().ToList();

			if (distinctGroup.Count > 0)
			{
				string combinedGroup = $"_{groups[i].rows[0].group}"; 
				for (int j = 0; j < groups[i].rows.Count; j++)
				{
					groups[i].rows[j].combinedGroup = combinedGroup;
				}
			}
		}
	}
	
	[HorizontalGroup("Group 7")]
	[InfoBox("Step 7")]
	[Button]
	public void CalculateGroupStatistic()
	{
		for (int i = 0; i < groups.Count; i++)
		{
			groups[i].CalculateStatistic();
		}
	}
	
	
	[HorizontalGroup("Group 7")]
	[InfoBox("Step 8")]
	[Button(ButtonSizes.Large)]
	public void GroupToSheet()
	{
		sheets[0].list.Clear();
		foreach (Group group in groups)
		{
			List<ParamModified> list = group.rows.Clone().ToList();
			
			for (int i = 0; i < list.Count; i++)
			{
				if (i == 0)
				{
					list[i].Description = group.name;
				}
				else if (i == 1)
				{
					list[i].Description = $"{group.description}";
				}
				// else if (i == 2)
				// {
				// 	list[i].Description = $"{group.comment}";
				// }
				else
				{
					list[i].Description = "";
				}

				sheets[0].list.Add(list[i]);
			}
		}
	}

	[Button]
	public void SortGroups()
	{
		groups = groups.OrderByDescending(x => x.machines).ThenByDescending(x => x.name).ToList();
		GroupToSheet();
	}

	public void DeleteComment()
	{
		for (int i = 0; i < groups.Count; i++)
		{
			groups[i].comment = "";
		}
	}

	public void RemoveGroup(string groupName)
	{
		if (string.IsNullOrEmpty(groupName))
		{
			Debug.LogWarning("Group name can not be null");
			return;
		}

		Group g = groups.Find(x => groupName.Equals(x.name));

		if (g == null)
		{
			Debug.LogWarning($"Couldn't find group {groupName} in {name}");
			return;
		}
		else
		{
			Debug.Log($"Found group {groupName} in {name}");
		}
		
		g.name = "";
		g.description = "";

		List<ParamModified> rows = sheets[0].list.FindAll(x => x.combinedGroup == g.rows[0].combinedGroup).ToList();

		for (int i = 0; i < rows.Count; i++)
		{
			rows[i].combinedGroup = "";
			rows[i].Description = "";

			g.rows[i].combinedGroup = "";
			g.rows[i].Description = "";
		}
		
		ClearGroups();
		SheetToGroup();
		SortGroups();
	}
	
	public void FillStatisticIntoExcelSheet(HSSFSheet sheet)
	{
		AssignCombinedGroupColumn();
		CalculateGroupStatistic();

		List<Group> list = groups;
		HSSFRow row = (HSSFRow) sheet.CreateRow(0);
		int i = 0;
		HSSFCell cell = (HSSFCell) row.CreateCell((short) i);
		cell.SetCellValue("Name");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Description");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Note");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Combined Group");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Unique Machines");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Unique UserID");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Count F");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Count H");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("EULA Machines");
		
		for (int j= 0; j < list.Count; j++)
		{
			Group param = list[j];
			HSSFRow r = (HSSFRow) sheet.CreateRow(j + 1);
			param.FillDataIntoRow(r);
		}
	}
	
	public void FillDataIntoExcelSheet(HSSFSheet sheet)
	{
		AssignCombinedGroupColumn();
		CalculateGroupStatistic();
		
		List<ParamModified> list = sheets[0].list;
		HSSFRow row = (HSSFRow) sheet.CreateRow(0);
		int i = 0;
		HSSFCell cell = (HSSFCell) row.CreateCell((short) i);
		cell.SetCellValue("Description");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Combined Group");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Group");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Data Rows Count");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Count Unique UserID");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Count Unique MachineID");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Count License F");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Count License H");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("IP");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("LicenseHash");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("MachineId");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Version");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Location");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("Userid");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("count");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("SerialNumber");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("LicenseType");
		cell = (HSSFCell) row.CreateCell((short) ++i);
		cell.SetCellValue("ShortVer");
		
		for (int j= 0; j < list.Count; j++)
		{
			ParamModified param = list[j];
			HSSFRow r = (HSSFRow) sheet.CreateRow( j + 1);
			param.FillDataIntoRow(r);
		}
	}
	
	[HorizontalGroup("Group 9")]
	[PropertyOrder(9)] 
	[InfoBox("Step 9: Export Group data to csv")]
	[FolderPath(AbsolutePath = true)]
	public string output_GroupToCSV;
	[HorizontalGroup("Group 9")]
	[PropertyOrder(9)]
	[Button(ButtonSizes.Gigantic)]
	public void ExportGroupToCSV()
	{
		string path = output_GroupToCSV;
		
		if (groups.Count == 0)
		{
			Debug.LogWarning("Groups is empty");
			return;
		}
		
		AssignCombinedGroupColumn();
		CalculateGroupStatistic();
		SortGroups();
		
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("Description,Combined Group,Group,Data Rows Count,Count Unique UserID,Count Unique MachineID,Count License F," +
		              "Count License H,IP,LicenseHash,MachineId,Version,Location,Userid,count,SerialNumber,LicenseType," +
		              "ShortVer");

		List<Group> orderedGroup = groups;

		foreach (Group group in orderedGroup)
		{
			List<ParamModified> list = group.rows.Clone().ToList();
			
			for (int i = 0; i < list.Count; i++)
			{
				if (i == 0)
				{
					list[i].Description = group.name;
				}
				else if (i == 1)
				{
					if (!string.IsNullOrEmpty(group.description))
					{
						list[i].Description = $"\"{group.description.Trim()}\"";
					}
				}
				// else if (i == 2)
				// {
				// 	list[i].Description = $"\" unique machines: {group.machines}\nuser: {group.users}\nF: {group.F}\nH: {group.H}\nrows: {group.rows.Count}\nEULA machines: {group.F + group.H}\"";
				// }
				else
				{
					list[i].Description = "";
				}
				
				sb.AppendLine(list[i].ToString());
			}
		}

		string finalPath = path.EndsWith($"{Path.DirectorySeparatorChar}") ? path : $"{path}{Path.DirectorySeparatorChar}";
		string finalFileName = $"{finalPath}{this.name}.csv";
		using (StreamWriter file = new StreamWriter(finalFileName))
		{
			file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
		}
        
		Debug.Log($"Final output: {finalFileName}");
	}

	[HorizontalGroup("Group 10")]
	[PropertyOrder(10)]
	[InfoBox("Step 10: Export Group Statistic data to csv")] 
	[FolderPath(AbsolutePath = true)]
	public string output_GroupStatisticToCSV;
	[HorizontalGroup("Group 10")]
	[PropertyOrder(10)] 
	[Button(ButtonSizes.Gigantic)]
	public void ExportGroupStatisticToCSV()
	{
		string path = output_GroupStatisticToCSV;
		if (groups.Count == 0)
		{
			Debug.LogWarning("Groups is empty");
			return;
		}
		
		AssignCombinedGroupColumn();
		CalculateGroupStatistic();
		SortGroups();
		
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("Name,Description,Note,Combined Group,Unique Machines,Unique UserID, Count F,Count H,EULA Machines");

		List<Group> orderedGroup = groups;
		foreach (Group group in orderedGroup)
		{
			sb.AppendLine($"{group.name},\"{group.description}\",\"{group.comment}\",{group.rows[0].combinedGroup},{group.machines},{group.users},{group.F},{group.H},{group.F + group.H}");
		}

		string finalPath = path.EndsWith($"{Path.DirectorySeparatorChar}") ? path : $"{path}{Path.DirectorySeparatorChar}";
		string finalFileName = $"{finalPath}{this.name}_statistic.csv";
		using (StreamWriter file = new StreamWriter(finalFileName))
		{
			file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
		}
		Debug.Log($"Final output: {finalFileName}");
	}

	void CrossCheck(List<Group> namedGroups1, string sheet1, List<Group> allGroupFromOtherSheet, string sheet2, float matchingThreshold = 0.2f, bool applyLabeling = false)
	{
		Debug.Log($"====================Find named groups of {sheet1} in {sheet2}=====================\n\n\n");
		StringBuilder log = new StringBuilder();
		for (int i = 0; i < namedGroups1.Count; i++)
		{
			for (int j = 0; j < allGroupFromOtherSheet.Count; j++)
			{
				Group g1 = namedGroups1[i];
				Group g2 = allGroupFromOtherSheet[j];

				
				if (g1.Compare(g2, sheet1, sheet2, log, matchingThreshold) > matchingThreshold)
				{
					if (applyLabeling)
					{
						if (string.IsNullOrEmpty(g2.name))
						{
							Debug.Log(log);
							
							g2.name = g1.name;
							g2.description = g1.description;
						}
						else
						{
							// string temp = $"[{sheet1}] [{g1.name}] vs [{g2.name}] [{sheet2}]";
							string temp = log.ToString();
							if (!g1.name.Equals(g2.name) && !g1.comment.Contains(temp))
							{
								g1.comment = $"{temp}\n{g1.comment}";
							}
							// Debug.LogWarning($"Won't set name for group 2 because it already has: g1 = [{g1.name}] \t \t g2 = [{g2.name}]");
						}
					}
					else
					{
						Debug.Log(log);
					}
				}
			}
		}
		Debug.Log($"==================================================================================\n\n\n");
	}

	[PropertyOrder(11)] public Entity_sheet_Modified referenceData;
	[PropertyOrder(11)]
	[Button]
	public void AutomateAllStep()
	{
		ClearGroups();
		SheetToGroup();
		if (referenceData != null)
		{
			MatchingNamedGroupsBetweenSheets(0.2f, true);
		}
		MergeDuplicatedGroupsByName();
		// ExportGroupToCSV();
		// ExportGroupStatisticToCSV();
	}
}
