using System.Collections;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using UnityEngine;
using Sirenix.OdinInspector;
using Sirenix.OdinInspector.Editor;
using Sirenix.Utilities;
using Sirenix.Utilities.Editor;

[CreateAssetMenu(menuName = "Tools/Matching Weeks")]
public class MatchingWeeks : ScriptableObject
{
    public int iterationCount = 3;
    [FolderPath(AbsolutePath = true)]
    public string outputPath;
    public Entity_sheet_Modified referenceData;
    public List<Entity_sheet_Modified> weeks;

    public string groupName;
    
    [Button]
    public void RemoveGroup()
    {
        for (int i = 0; i < weeks.Count; i++)
        {
            weeks[i].RemoveGroup(groupName.Trim());
        }
    }
    
    [Button]
    public void DeleteComment()
    {
        for (int i = 0; i < weeks.Count; i++)
        {
            weeks[i].DeleteComment();
        }
    }
    
    [Button]
    public void TrimGroupName()
    {
        for (int i = 0; i < weeks.Count; i++)
        {
            weeks[i].TrimGroupName();
        }
    }
    
    [Button]
    public void CompareAndMatchWeeks() {

        int k = 0, countMerge = 0;
        do
        {
            if(k == 0) {
                countMerge = 1;
            } 
            else {
                countMerge = 0;
            }

            for (int i = 0; i < weeks.Count; i++)
            {
                weeks[i].referenceData = referenceData;
                countMerge += weeks[i].MatchingNamedGroupsBetweenSheets(applyLabeling : true);
            }
            ++k;
        }while(k < iterationCount && countMerge > 0);

        Debug.Log($"iteration: {k} \t countMerge: {countMerge}");
    }

    [Button]
    public void ExportData() {
        List<Entity_sheet_Modified> list = new List<Entity_sheet_Modified>();
        // list.Add(referenceData);
        list.AddRange(weeks);

        string finalPath = outputPath.EndsWith($"{Path.DirectorySeparatorChar}") ? outputPath : $"{outputPath}{Path.DirectorySeparatorChar}";
        string fileNameStatistic = $"{finalPath}{this.name}_statistic.xls";
        string fileNameData = $"{finalPath}{this.name}_data.xls";

        if (File.Exists(fileNameData))
        {
            File.Delete(fileNameData);
        }
        
        if (File.Exists(fileNameStatistic))
        {
            File.Delete(fileNameStatistic);
        }
        
        FileStream fileData = new FileStream(fileNameData, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
        FileStream fileDStat = new FileStream(fileNameStatistic, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
        
        HSSFWorkbook bookStat = new HSSFWorkbook();
        HSSFWorkbook bookData = new HSSFWorkbook();
        
        
        for (int i = 0; i < list.Count; i++)
        {
            HSSFSheet sheetStat = (HSSFSheet)bookStat.CreateSheet(list[i].name);
            HSSFSheet sheetData = (HSSFSheet)bookData.CreateSheet(list[i].name);
            
            list[i].SortGroups();
            list[i].FillDataIntoExcelSheet(sheetData);
            list[i].FillStatisticIntoExcelSheet(sheetStat);
            
            /*temp = list[i].output_GroupStatisticToCSV;
            temp1 = list[i].output_GroupToCSV;
            list[i].output_GroupStatisticToCSV = outputPath;
            list[i].output_GroupToCSV = outputPath;
            list[i].ExportGroupStatisticToCSV();
            list[i].ExportGroupToCSV();
            list[i].output_GroupStatisticToCSV = temp;
            list[i].output_GroupToCSV = temp1;*/
        }
        
        bookData.Write(fileData);
        bookData.Close();
        fileData.Close();
        Debug.Log($"{fileNameData}");
        
        bookStat.Write(fileDStat);
        bookStat.Close();
        fileDStat.Close();
        Debug.Log($"{fileNameStatistic}");
        
    }

    
}
